"""
Timesheets - hosted (multi-company) web server.

Wraps the desktop app's logic (timesheet_app.py) so it runs as a normal web
app: a company signs up online, gets its own isolated SQLite database, and its
staff sign in from any browser. The desktop file dialogs become browser
uploads/downloads. Nothing in timesheet_app.py is changed - this file sits
around it.

Run:   uvicorn server:app --host 0.0.0.0 --port 8000
Env:   DATA_ROOT   where company data lives (default ./data)
       SECRET_KEY  signs session cookies (set a long random one in production)
       SECURE_COOKIES=1 behind HTTPS (Render/Fly/etc.)
"""
import io, os, sys, json, hmac, hashlib, base64, time, re, secrets, shutil, threading, urllib.parse
import datetime as dt
import types

# ---------------------------------------------------------------------------
# A stand-in for the pywebview module: the app only uses its dialog constants.
# ---------------------------------------------------------------------------
_wv = types.ModuleType("webview")
_wv.OPEN_DIALOG, _wv.FOLDER_DIALOG, _wv.SAVE_DIALOG = 10, 20, 30
sys.modules.setdefault("webview", _wv)

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))          # timesheet_app.py lives one level up
DATA_ROOT = os.path.abspath(os.environ.get("DATA_ROOT") or os.path.join(HERE, "data"))
os.makedirs(DATA_ROOT, exist_ok=True)
# Keep the desktop module's global folders inside DATA_ROOT and never put a
# logo there (a logo at DATA_DIR would otherwise show for every company).
os.environ.setdefault("APPDATA", os.path.join(DATA_ROOT, "_global"))
os.environ.setdefault("LOCALAPPDATA", os.path.join(DATA_ROOT, "_global"))

import timesheet_app as ts                           # noqa: E402

from fastapi import FastAPI, Request, Response, HTTPException   # noqa: E402
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse, RedirectResponse  # noqa: E402
from starlette.concurrency import run_in_threadpool  # noqa: E402
import sqlite3                                       # noqa: E402

SECRET = (os.environ.get("SECRET_KEY") or "").encode() or None
if not SECRET:
    # Dev convenience: a per-install secret written once to DATA_ROOT.
    sp = os.path.join(DATA_ROOT, "_global", "secret.key")
    os.makedirs(os.path.dirname(sp), exist_ok=True)
    if not os.path.exists(sp):
        with open(sp, "w") as f:
            f.write(secrets.token_hex(32))
    SECRET = open(sp).read().strip().encode()
SECURE_COOKIES = os.environ.get("SECURE_COOKIES", "0") == "1"
COMPANIES_DIR = os.path.join(DATA_ROOT, "companies")
EXPORTS_DIR = os.path.join(DATA_ROOT, "_exports")
os.makedirs(COMPANIES_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)
ACCOUNTS_DB = os.path.join(DATA_ROOT, "accounts.db")


# ---------------------------------------------------------------------------
# Accounts index: which company an email belongs to
# ---------------------------------------------------------------------------
def _acc():
    con = sqlite3.connect(ACCOUNTS_DB, timeout=30)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA busy_timeout=15000")
    return con


with _acc() as con:
    con.executescript("""
    CREATE TABLE IF NOT EXISTS companies(
        id INTEGER PRIMARY KEY, name TEXT NOT NULL, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS directory(
        email TEXT NOT NULL, company_id INTEGER NOT NULL, user_id INTEGER NOT NULL,
        PRIMARY KEY(email, company_id));
    CREATE INDEX IF NOT EXISTS ix_dir_email ON directory(email);
    """)


def company_dir(cid):
    return os.path.join(COMPANIES_DIR, str(int(cid)))


def company_cfg_path(cid):
    return os.path.join(company_dir(cid), "config.json")


def load_company_cfg(cid):
    cfg = dict(ts.DEFAULT_CONFIG)
    cfg["db_path"] = os.path.join(company_dir(cid), "timesheets.db")
    try:
        with open(company_cfg_path(cid), "r", encoding="utf-8") as fh:
            cfg.update(json.load(fh))
    except Exception:
        pass
    cfg["db_path"] = os.path.join(company_dir(cid), "timesheets.db")   # never movable online
    cfg["_cid"] = int(cid)
    return cfg


def _save_company_cfg(cfg):
    cid = cfg.get("_cid")
    if cid is None:
        return
    out = {k: v for k, v in cfg.items() if not k.startswith("_") and k != "db_path"}
    os.makedirs(company_dir(cid), exist_ok=True)
    with open(company_cfg_path(cid), "w", encoding="utf-8") as fh:
        json.dump(out, fh, indent=2)


ts.save_config = _save_company_cfg        # the app's settings code calls this by name

_DBS = {}
_DBS_LOCK = threading.Lock()


def company_db(cid):
    cid = int(cid)
    with _DBS_LOCK:
        db = _DBS.get(cid)
        if db is None:
            db = ts.DB(os.path.join(company_dir(cid), "timesheets.db"))
            db.seed_defaults()
            _DBS[cid] = db
        return db


def sync_directory(cid):
    """Mirror a company's user emails into the accounts index."""
    db = company_db(cid)
    users = db.list_users(include_inactive=True)
    with _acc() as con:
        con.execute("DELETE FROM directory WHERE company_id=?", (cid,))
        con.executemany("INSERT OR REPLACE INTO directory(email, company_id, user_id) VALUES(?,?,?)",
                        [((u["email"] or "").strip().lower(), cid, u["id"]) for u in users if u["email"]])


# ---------------------------------------------------------------------------
# Sessions: signed cookie {cid, uid, exp}
# ---------------------------------------------------------------------------
COOKIE = "ts_session"


def _sign(payload: bytes) -> str:
    sig = hmac.new(SECRET, payload, hashlib.sha256).digest()
    return base64.urlsafe_b64encode(payload).decode() + "." + base64.urlsafe_b64encode(sig).decode()


def make_session(cid, uid, days, email=""):
    exp = int(time.time()) + int(days * 86400)
    return _sign(json.dumps({"cid": int(cid), "uid": int(uid), "em": (email or "").lower(), "exp": exp}).encode())


def read_session(token):
    try:
        p, s = token.split(".", 1)
        payload = base64.urlsafe_b64decode(p.encode())
        sig = base64.urlsafe_b64decode(s.encode())
        if not hmac.compare_digest(sig, hmac.new(SECRET, payload, hashlib.sha256).digest()):
            return None
        d = json.loads(payload)
        if d.get("exp", 0) < time.time():
            return None
        return d
    except Exception:
        return None


def file_token(cid, path, minutes=15):
    exp = int(time.time()) + minutes * 60
    return _sign(json.dumps({"cid": int(cid), "p": path, "exp": exp}).encode())


def read_file_token(tok):
    try:
        p, sg = tok.split(".", 1)
        payload = base64.urlsafe_b64decode(p.encode())
        sig = base64.urlsafe_b64decode(sg.encode())
        if not hmac.compare_digest(sig, hmac.new(SECRET, payload, hashlib.sha256).digest()):
            return None
        d = json.loads(payload)
        return d if d.get("exp", 0) >= time.time() else None
    except Exception:
        return None


def set_cookie(resp: Response, token, days):
    resp.set_cookie(COOKIE, token, max_age=int(days * 86400), httponly=True,
                    samesite="lax", secure=SECURE_COOKIES, path="/")


def clear_cookie(resp: Response):
    resp.delete_cookie(COOKIE, path="/")


# ---------------------------------------------------------------------------
# The per-request API object
# ---------------------------------------------------------------------------
class FakeWindow:
    """Replaces pywebview's window: 'save' dialogs hand back a path in the
    exports folder (served as a download afterwards); 'open' dialogs hand back
    the file the browser uploaded for this request."""
    def __init__(self, api):
        self.api = api

    def create_file_dialog(self, kind, directory="", allow_multiple=False,
                           save_filename="", file_types=()):
        if kind == _wv.SAVE_DIALOG:
            safe = re.sub(r"[^A-Za-z0-9._ -]", "_", save_filename or "export")[:90]
            d = os.path.join(EXPORTS_DIR, str(self.api.cid))
            os.makedirs(d, exist_ok=True)
            return os.path.join(d, f"{secrets.token_hex(6)}_{safe}")
        if kind == _wv.OPEN_DIALOG:
            return self.api._upload or None
        return None


# Role gates. The desktop app hid these in the UI; online they must be enforced.
MANAGER_ONLY = {"approvals", "approve", "approval_lines", "save_admin_lines", "job_costs",
                "rates_for_job", "save_rates", "run_report", "export_report", "job_cost_pdf",
                "month_end_export", "export_timesheet_template", "import_external_timesheet",
                "analytics", "job_timeline", "job_timeline_detail", "timeline_lines", "job_billing",
                "job_documents", "document_folders", "dash_jobs_at_risk", "export_invoice_template",
                "import_invoices", "scan_document", "ocr_status"}
ADMIN_ONLY = {"bug_reports","set_bug_status","audit", "users", "jobs", "roles", "add_user", "update_user", "reset_password",
              "delete_user", "add_job", "update_job", "delete_job", "add_role", "toggle_role",
              "tidy_documents", "get_settings", "save_settings", "set_logo", "remove_logo",
              "save_goals", "add_variation", "update_variation", "delete_variation", "add_invoice",
              "update_invoice", "set_invoice_paid", "delete_invoice", "add_bill", "update_bill",
              "set_bill_paid", "delete_bill", "add_document", "delete_document", "attach_payment"}

# Methods that only make sense on the desktop, or that this server replaces.
BLOCKED = {"pick_data_folder", "complete_setup", "change_data_folder",
           "open_data_folder", "open_file", "open_folder", "pick_file",
           "pick_document", "pick_logo", "startup_prefs", "remember_login",
           "auto_login", "forget_login", "login", "logout", "session",
           "needs_setup"}


class WebApi(ts.Api):
    def __init__(self, cid=None, user=None, upload=None):
        self.cid = int(cid) if cid is not None else None
        self.cfg = load_company_cfg(cid) if cid is not None else dict(ts.DEFAULT_CONFIG)
        self.db = company_db(cid) if cid is not None else None
        self.user = user
        self.window = FakeWindow(self)
        self._upload = upload

    # -- things the desktop stored per PC --
    def brand(self):
        return {"company": (self.cfg.get("company_name") or "").strip() if self.cid else "",
                "app": ts.APP_NAME, "version": ts.APP_VERSION,
                "logo": self.logo() if self.cid else "", "needs_setup": False,
                "online": True}

    def report_bug(self, details, steps="", page="", device="", errors=""):
        r = super().report_bug(details, steps, page, device, errors)
        if r.get("ok"):
            # Also keep one log across every company so the product owner sees them all.
            try:
                rec = {"time": dt.datetime.now().isoformat(timespec="seconds"), "company_id": self.cid,
                       "company": (self.cfg.get("company_name") or self.db.get_meta("company_name") or ""),
                       "id": r["id"], "user": self.user.get("email"), "page": page,
                       "details": details, "steps": steps, "device": device, "errors": errors}
                with open(os.path.join(DATA_ROOT, "bug_reports.jsonl"), "a", encoding="utf-8") as f:
                    f.write(json.dumps(rec) + "\n")
            except Exception:
                pass
        return r

    def app_config(self):
        self._need()
        return {"timeout_minutes": int(self.cfg.get("inactivity_timeout_minutes", 15)),
                "daily_warning": float(self.cfg.get("daily_hours_warning", 10)),
                "db_path": ""}

    def get_settings(self):
        s = super().get_settings()
        s["online"] = True
        s["data_folder"] = ""
        return s

    def set_logo(self, path):
        """Browser upload (already saved via save_dropped_file) -> company logo."""
        u = self._admin()
        path = os.path.abspath(path or "")
        if not path.startswith(os.path.abspath(company_dir(self.cid))) or not os.path.isfile(path):
            return {"ok": False, "error": "Upload the logo again."}
        if os.path.getsize(path) > 700_000:
            os.remove(path)
            return {"ok": False, "error": "Logo must be under 700 KB. Resize it and try again."}
        with open(path, "rb") as f:
            head = f.read(8)
        if head[:8] != b"\x89PNG\r\n\x1a\n":
            os.remove(path)
            return {"ok": False, "error": "The logo must be a PNG file."}
        shutil.move(path, os.path.join(company_dir(self.cid), "logo.png"))
        self._logo_cache = ts._logo_data_uri(self.cfg)
        self.db.audit(u["email"], "SETTINGS", "", "Logo changed")
        return {"ok": True, "logo": self._logo_cache}

    def file_link(self, path):
        """A link to open a document or export that works without the session
        cookie (desktop app hands it to the system browser). Valid 15 minutes."""
        self._need()
        path = os.path.abspath(path or "")
        allowed = (os.path.abspath(company_dir(self.cid)) + os.sep,
                   os.path.abspath(os.path.join(EXPORTS_DIR, str(self.cid))) + os.sep)
        if not path.startswith(allowed) or not os.path.isfile(path):
            return {"ok": False, "error": "That file isn't here any more."}
        return {"ok": True, "url": "/api/open?t=" + urllib.parse.quote(file_token(self.cid, path))}

    def remove_logo(self):
        self._admin()
        try:
            os.remove(os.path.join(company_dir(self.cid), "logo.png"))
        except FileNotFoundError:
            pass
        self._logo_cache = ""
        return {"ok": True}

    # user changes must keep the sign-in directory current
    def add_user(self, *a, **k):
        r = super().add_user(*a, **k); sync_directory(self.cid); return r

    def update_user(self, *a, **k):
        r = super().update_user(*a, **k); sync_directory(self.cid); return r

    def delete_user(self, *a, **k):
        r = super().delete_user(*a, **k); sync_directory(self.cid); return r


def _public_user(row):
    u = dict(row)
    for k in ("password_hash", "salt", "auth_token"):
        u.pop(k, None)
    return u


_FAILS = {}
_FAILS_LOCK = threading.Lock()


def _too_many(key, limit=10, window=900):
    now = time.time()
    with _FAILS_LOCK:
        hits = [t for t in _FAILS.get(key, []) if now - t < window]
        _FAILS[key] = hits
        return len(hits) >= limit


def _record_fail(key):
    with _FAILS_LOCK:
        _FAILS.setdefault(key, []).append(time.time())


def _clear_fail(key):
    with _FAILS_LOCK:
        _FAILS.pop(key, None)


def find_login(ident, password):
    """Email (or the part before @) + password -> (cid, user) or None."""
    ident = (ident or "").strip().lower()
    if not ident:
        return None
    with _acc() as con:
        if "@" in ident:
            rows = con.execute("SELECT company_id FROM directory WHERE email=?", (ident,)).fetchall()
        else:
            rows = con.execute("SELECT company_id FROM directory WHERE email LIKE ?",
                               (ident + "@%",)).fetchall()
    for r in rows:
        try:
            u = company_db(r["company_id"]).authenticate(ident, password)
        except Exception:
            u = None
        if u:
            return r["company_id"], _public_user(u)
    return None


# ---------------------------------------------------------------------------
# The page: the desktop HTML with the pywebview bridge replaced by HTTP
# ---------------------------------------------------------------------------
BRIDGE_JS = r"""<script>
(function(){
  async function call(name,args,upload){
    const body={args:args||[]}; if(upload) body.upload=upload;
    let r; try{ r=await fetch('/api/'+name,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body),credentials:'same-origin'}); }
    catch(e){ throw new Error('No connection to the server. Check your internet and try again.'); }
    if(r.status===401){ if(window.USER){ window.USER=null; location.reload(); } return null; }
    let j=null; try{ j=await r.json(); }catch(e){ throw new Error('Server error ('+r.status+')'); }
    if(j&&j.error){ throw new Error(j.error); }
    return j?j.result:null;
  }
  function pickFile(accept){ return new Promise(res=>{ const i=document.createElement('input'); i.type='file'; if(accept)i.accept=accept; i.style.display='none'; document.body.appendChild(i);
    let done=false; const finish=v=>{ if(done)return; done=true; i.remove(); res(v); };
    i.onchange=()=>finish(i.files&&i.files[0]||null); window.addEventListener('focus',()=>setTimeout(()=>{ if(!i.files||!i.files.length) finish(null); },800),{once:true}); i.click(); }); }
  function toB64(file){ return new Promise((res,rej)=>{ const fr=new FileReader(); fr.onload=()=>res(String(fr.result).split(',')[1]||''); fr.onerror=()=>rej(new Error('Could not read the file')); fr.readAsDataURL(file); }); }
  async function upload(accept){ const f=await pickFile(accept); if(!f) return null; if(f.size>25*1024*1024){ alert('That file is too big (25 MB limit).'); return null; }
    const r=await call('save_dropped_file',[f.name,await toB64(f)]); if(!r||!r.ok){ alert((r&&r.error)||'Upload failed'); return null; } return r.path; }
  const CLIENT=(function(){ try{ if(location.search.indexOf('client=1')>=0) localStorage.setItem('ts_client','1'); return localStorage.getItem('ts_client')==='1'; }catch(e){ return location.search.indexOf('client=1')>=0; } })();
  window.TS_CLIENT=CLIENT;
  async function openPath(p){ if(!p) return false;
    let r=null; try{ r=await call('file_link',[p]); }catch(e){}
    const u=(r&&r.ok)?r.url:('/api/open?p='+encodeURIComponent(p));
    if(r&&!r.ok&&r.error){ alert(r.error); return false; }
    const a=document.createElement('a'); a.href=location.origin+u; a.target='_blank'; a.rel='noopener'; document.body.appendChild(a); a.click(); setTimeout(()=>a.remove(),500); return true; }
  const special={
    pick_file:()=>upload('.pdf,.png,.jpg,.jpeg,.gif,.heic'),
    pick_document:()=>upload('.pdf,.png,.jpg,.jpeg,.tif,.tiff'),
    import_external_timesheet:async()=>{ const p=await upload('.xlsx'); if(!p) return {ok:false}; return call('import_external_timesheet',[],p); },
    import_invoices:async()=>{ const p=await upload('.xlsx'); if(!p) return {ok:false}; return call('import_invoices',[],p); },
    pick_logo:async()=>{ const p=await upload('.png'); if(!p) return {ok:false}; return call('set_logo',[p]); },
    open_file:async(p)=>openPath(p),
    open_folder:async()=>{ alert('Documents are stored online. Use the open buttons on each file instead.'); return false; },
    open_data_folder:async()=>false, pick_data_folder:async()=>({ok:false}),
    startup_prefs:async()=>({username:(localStorage.getItem('ts_user')||'')}),
    remember_login:async(rem,username)=>{ try{ if(username) localStorage.setItem('ts_user',username); }catch(e){} return call('remember_login',[!!rem]); },
    auto_login:async()=>call('auto_login',[]),
    forget_login:async()=>{ return call('forget_login',[]); },
  };
  const api=new Proxy({},{get:(t,name)=>{ if(typeof name!=='string'||name==='then') return undefined; if(special[name]) return special[name]; return (...args)=>call(name,args); }});
  try{ Object.defineProperty(window,'pywebview',{value:{api:api},writable:false,configurable:false}); }catch(e){ window.pywebview={api:api}; }
  window.addEventListener('DOMContentLoaded',()=>{ try{ window.dispatchEvent(new Event('pywebviewready')); }catch(e){} });
})();
</script>"""

IMPORT_JS = r"""function _pickRaw(accept){ return new Promise(res=>{ const i=document.createElement('input'); i.type='file'; i.accept=accept; i.style.display='none'; document.body.appendChild(i); i.onchange=()=>{const f=i.files&&i.files[0]||null; i.remove(); res(f);}; i.click(); }); }
async function _importUpload(url,file,msg){
  const m=E('imp_msg'); m.textContent=msg; const fd=new FormData(); fd.append('file',file,file.name);
  let r; try{ r=await fetch(url,{method:'POST',body:fd,credentials:'same-origin'}); }catch(e){ m.textContent='Upload failed: '+e.message; return null; }
  let j=null; try{ j=await r.json(); }catch(e){}
  if(!r.ok||!j||!j.ok){ m.textContent='Import failed: '+((j&&j.error)||('server error '+r.status)); return null; }
  return j;
}
async function importDb(){
  const f=await _pickRaw('.db,.sqlite,.sqlite3'); if(!f)return;
  if(!confirm('Replace EVERYTHING in this company with the contents of '+f.name+'?\n\nPeople, timesheets, jobs, invoices, bills and expenses currently here will be replaced. A copy is kept on the server.'))return;
  const j=await _importUpload('/api/import_database',f,'Uploading and checking the database…'); if(!j)return;
  E('imp_msg').textContent=`Imported: ${j.users} people, ${j.timesheets} timesheets, ${j.jobs} jobs, ${j.expenses} expenses.`;
  alert('Database imported.\n\n'+j.users+' people, '+j.timesheets+' timesheets, '+j.jobs+' jobs.\n\nYou will now be signed out. Sign back in with your old desktop login, then upload the Receipts zip from Settings (step 2).');
  USER=null; location.reload();
}
async function importReceipts(){
  const f=await _pickRaw('.zip'); if(!f)return;
  const j=await _importUpload('/api/import_receipts',f,'Uploading documents ('+Math.round(f.size/1048576)+' MB)…'); if(!j)return;
  E('imp_msg').textContent=`Documents imported: ${j.files} files. ${j.relinked} attachment links updated.`;
}
"""


INSTALL_JS = r"""<style>
.inst{position:fixed;inset:0;z-index:90;background:#0f1117;color:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:calc(28px + env(safe-area-inset-top)) 26px calc(22px + env(safe-area-inset-bottom));animation:instIn .35s ease both;overflow:auto;background-image:radial-gradient(ellipse at 50% 0%,rgba(108,192,74,.22) 0,transparent 60%)}
@keyframes instIn{from{opacity:0}to{opacity:1}}
.inst .ic{width:96px;height:96px;border-radius:24px;background:#fff;box-shadow:0 14px 40px rgba(0,0,0,.5);margin-bottom:22px;animation:instPop .6s .1s cubic-bezier(.34,1.4,.64,1) both}
@keyframes instPop{from{transform:scale(.6);opacity:0}to{transform:none;opacity:1}}
.inst h2{font-size:26px;font-weight:800;letter-spacing:-.02em;margin:0 0 8px;line-height:1.15}
.inst .lead{color:#aab1be;font-size:15px;line-height:1.5;margin:0 0 26px;max-width:320px}
.inst ol{list-style:none;padding:0;margin:0 0 26px;width:100%;max-width:340px;text-align:left;counter-reset:st}
.inst li{display:flex;align-items:center;gap:14px;background:#1a1d26;border:1px solid #262b36;border-radius:14px;padding:14px 16px;margin-bottom:10px;font-size:15px;line-height:1.35;counter-increment:st;animation:instUp .45s cubic-bezier(.22,1,.36,1) both}
.inst li:nth-child(2){animation-delay:.08s}.inst li:nth-child(3){animation-delay:.16s}
@keyframes instUp{from{transform:translateY(14px);opacity:0}to{transform:none;opacity:1}}
.inst li::before{content:counter(st);flex:0 0 30px;height:30px;border-radius:50%;background:#6cc04a;color:#0e1a08;font-weight:800;display:flex;align-items:center;justify-content:center;font-size:14px}
.inst li b{color:#fff}
.inst .chip{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;border-radius:8px;background:#2a2f3a;vertical-align:middle;margin:0 2px}
.inst .go{width:100%;max-width:340px;background:#6cc04a;color:#0e1a08;border:0;border-radius:14px;padding:17px;font-weight:800;font-size:17px;cursor:pointer;margin-bottom:10px}
.inst .later{background:none;border:0;color:#7d8494;font-size:14px;padding:12px;cursor:pointer;margin-top:auto}
.inst .arrow{position:fixed;left:50%;bottom:calc(10px + env(safe-area-inset-bottom));transform:translateX(-50%);color:#6cc04a;animation:instBob 1.4s ease-in-out infinite}
@keyframes instBob{0%,100%{transform:translate(-50%,0)}50%{transform:translate(-50%,8px)}}
.inst .spacer{flex:1 1 auto;min-height:10px}
</style><script>
(function(){
  var ua=navigator.userAgent||'';
  var isIOS=/iPhone|iPad|iPod/.test(ua)||(navigator.platform==='MacIntel'&&navigator.maxTouchPoints>1);
  var isAndroid=/Android/.test(ua);
  var installed=(window.matchMedia&&matchMedia('(display-mode: standalone)').matches)||navigator.standalone===true;
  var deferred=null;
  window.addEventListener('beforeinstallprompt',function(e){e.preventDefault();deferred=e;});
  window.addEventListener('appinstalled',function(){var b=document.getElementById('inst');if(b)b.remove();});
  if('serviceWorker' in navigator){navigator.serviceWorker.register('/sw.js').catch(function(){});}
  function dismissedRecently(){try{var t=+localStorage.getItem('ts_inst_x')||0;return Date.now()-t<3*864e5;}catch(e){return false;}}
  window.maybeInstallBanner=function(){
    if(installed||(!isIOS&&!isAndroid)||dismissedRecently()||document.getElementById('inst'))return;
    try{if(localStorage.getItem('ts_client'))return;}catch(e){}
    setTimeout(function(){
      var b=document.createElement('div');b.className='inst';b.id='inst';
      var share='<span class="chip"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12M7 8l5-5 5 5"/><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"/></svg></span>';
      var plus='<span class="chip"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round"><rect x="3" y="3" width="18" height="18" rx="4"/><path d="M12 8v8M8 12h8"/></svg></span>';
      var html='<div class="spacer"></div><img class="ic" src="/icon.png" alt="">';
      if(isIOS){
        html+='<h2>Add Timesheets to your Home Screen</h2><p class="lead">It opens full screen like a normal app, and you stay signed in.</p>'
          +'<ol><li><span>Tap '+share+' <b>Share</b> at the bottom of Safari <span style="color:#7d8494">(or under ⋯)</span></span></li>'
          +'<li><span>Scroll down and tap '+plus+' <b>Add to Home Screen</b></span></li>'
          +'<li><span>Tap <b>Add</b> — then open Timesheets from your home screen</span></li></ol>'
          +'<div class="spacer"></div><button class="later" id="instX">Not now</button>'
          +'<svg class="arrow" width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M6 13l6 6 6-6"/></svg>';
      }else{
        html+='<h2>Install the Timesheets app</h2><p class="lead">One tap. It goes on your home screen, opens full screen, and you stay signed in.</p>'
          +'<button class="go" id="instGo">Install app</button><p class="lead" id="instHelp" style="display:none"></p>'
          +'<div class="spacer"></div><button class="later" id="instX">Not now</button>';
      }
      b.innerHTML=html; document.body.appendChild(b);
      document.getElementById('instX').onclick=function(){try{localStorage.setItem('ts_inst_x',Date.now());}catch(e){} b.remove();};
      var go=document.getElementById('instGo');
      if(go)go.onclick=async function(){
        if(deferred){deferred.prompt();try{var r=await deferred.userChoice;if(r&&r.outcome==='accepted')b.remove();}catch(e){} deferred=null;}
        else{var h=document.getElementById('instHelp');h.style.display='block';h.innerHTML='Tap the browser menu <b style="color:#fff">⋮</b> (top right), then <b style="color:#fff">Install app</b> or <b style="color:#fff">Add to Home screen</b>.';go.style.display='none';}
      };
    },900);
  };
})();
</script>"""


def build_index_html():
    html = ts.INDEX_HTML.replace("{{LOGO}}", "").replace("{{LOGOCLS}}", "")

    def rep(old, new, count=1):
        nonlocal html
        assert html.count(old) == count, f"page template changed: {old[:80]!r}"
        html = html.replace(old, new)

    # bridge first
    rep("<script>", BRIDGE_JS + INSTALL_JS + "\n<script>")
    # setup card -> online sign-up card (no data-folder choice)
    rep('''<div class="brand">Welcome</div><div class="brand-sub">SET UP TIMESHEETS</div>''',
        '''<div class="brand">Create your company</div><div class="brand-sub">FREE TO START</div>''')
    rep('''    <label class="muted s-lab">Where should the data live?</label>''',
        '''    <div class="hidden"><label class="muted s-lab">Where should the data live?</label>''')
    rep('''    <div class="s-hint" id="su_hint">Settings, database and receipts are kept in your AppData folder. You can move them to a shared folder later in Settings.</div>''',
        '''    <div class="s-hint" id="su_hint"></div></div>
    <div class="s-hint">You'll be the admin. Add your team under Users once you're in.</div>''')
    rep('''    <button class="btn" id="sbtn" onclick="doSetup()">Create company &amp; sign in</button>''',
        '''    <button class="btn" id="sbtn" onclick="doSetup()">Create company &amp; sign in</button>
    <div class="s-hint" style="text-align:center;margin-top:12px"><a href="#" onclick="showLogin();return false;">Already have an account? Sign in</a></div>''')
    rep('''  <button class="btn" id="lbtn">Sign in</button>
</div></div></div>''', '''  <button class="btn" id="lbtn">Sign in</button>
  <div class="s-hint" style="text-align:center;margin-top:14px"><a href="#" onclick="showSignup();return false;">New here? Create your company</a></div>
</div></div></div>''')
    # sign-up flow talks to /api/signup and lands straight in the app
    rep('''async function doSetup(){''', '''function showSignup(){ E('login').classList.add('hidden'); E('setup').classList.remove('hidden'); setTimeout(()=>{const c=E('su_co');if(c&&(!document.activeElement||document.activeElement===document.body))c.focus();},50); }
function showLogin(){ E('setup').classList.add('hidden'); E('login').classList.remove('hidden'); }
async function doSetup(){
  const btn=E('sbtn'); if(btn.dataset.busy==='1')return;
  E('serr').textContent=''; btn.dataset.busy='1'; btn.textContent='Setting up\\u2026';
  let r=null; try{ r=await API.signup(E('su_co').value,E('su_name').value,E('su_email').value,E('su_pw').value); }catch(e){ r={ok:false,error:String(e&&e.message||e)}; }
  btn.dataset.busy=''; btn.textContent='Create company & sign in';
  if(!r||!r.ok){E('serr').textContent=(r&&r.error)||'Sign-up failed.';return;}
  try{ localStorage.setItem('ts_user',E('su_email').value.trim()); }catch(e){}
  E('setup').classList.add('hidden'); await applyBrand(); enterApp(r.user);
}
async function _desktopSetup(){''')
    # brand is per company: refresh it once signed in
    rep('''async function enterApp(u){
  USER=u;''', '''async function enterApp(u){
  USER=u; await applyBrand(); if(window.maybeInstallBanner) maybeInstallBanner();''')
    # settings: no data-folder card online
    rep('''  <div class="card" style="margin-top:16px"><div style="font-weight:700;margin-bottom:6px">Data</div>''',
        '''  <div class="card" style="margin-top:16px"><div style="font-weight:700;margin-bottom:6px">Your data</div>
    <div class="muted" style="font-size:13px;line-height:1.5">Stored online in your company's own private database, backed up daily.</div>
    <div class="muted" style="font-size:11.5px;margin-top:12px">${esc(window._brand?window._brand.app:'Timesheets')} ${esc(s.version)}</div></div>
  <div class="card" style="margin-top:16px"><div style="font-weight:700;margin-bottom:6px">Import from the desktop version</div>
    <div class="muted" style="font-size:13px;line-height:1.5">Moving from the old desktop app? Upload its <b>timesheets.db</b> file and everything comes across: people (with their passwords), timesheets, jobs, rates, invoices, bills, expenses and the audit trail. Then upload the <b>Receipts</b> folder as a zip so attached documents open here too.</div>
    <div style="display:flex;gap:8px;margin-top:12px;flex-wrap:wrap"><button class="btn sm" onclick="importDb()">1. Upload timesheets.db…</button><button class="btn sm ghost" onclick="importReceipts()">2. Upload Receipts.zip…</button></div>
    <div class="muted" style="font-size:12px;margin-top:10px;line-height:1.5"><b>This replaces everything currently in this company</b> — do it on a fresh company before anyone starts using it. A copy of the current data is kept on the server. Afterwards you sign in with your old desktop login.</div>
    <div id="imp_msg" class="muted" style="font-size:13px;margin-top:8px"></div></div>
  <div class="hidden"><div style="font-weight:700;margin-bottom:6px">Data</div>''')
    rep('''async function saveSettings(){''', IMPORT_JS + '''
async function saveSettings(){''')
    rep('''    <div class="k">Open maximised</div>''', '''    <div class="k hidden">Open maximised</div>''')
    rep('''<label style="font-size:13px"><input type="checkbox" id="st_max"''', '''<label class="hidden" style="font-size:13px"><input type="checkbox" id="st_max"''')
    rep('<meta charset="utf-8">',
        '<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Timesheets</title>'
        '<link rel="manifest" href="/manifest.webmanifest"><link rel="icon" href="/icon.png"><link rel="apple-touch-icon" href="/icon.png">'
        '<meta name="theme-color" content="#15151c"><meta name="apple-mobile-web-app-capable" content="yes">'
        '<meta name="mobile-web-app-capable" content="yes"><meta name="apple-mobile-web-app-title" content="Timesheets">'
        '<meta name="apple-mobile-web-app-status-bar-style" content="black">')
    return html


INDEX_HTML = build_index_html()

# ---------------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------------
app = FastAPI(title="Timesheets", docs_url=None, redoc_url=None)


def current(request: Request):
    tok = request.cookies.get(COOKIE)
    if not tok:
        return None
    s = read_session(tok)
    if not s:
        return None
    try:
        row = company_db(s["cid"]).get_user(int(s["uid"]))
    except Exception:
        return None
    if not row or not row["active"]:
        return None
    if (row["email"] or "").lower() != s.get("em", ""):
        return None          # account replaced (e.g. after an import): session no longer valid
    return s["cid"], _public_user(row)


@app.get("/", response_class=HTMLResponse)
def index():
    return HTMLResponse(INDEX_HTML)


ICON_PATH = os.path.join(os.path.dirname(HERE), "timesheets.png")


@app.get("/manifest.webmanifest")
def manifest():
    return JSONResponse({"name": "Timesheets", "short_name": "Timesheets", "start_url": "/",
                         "display": "standalone", "background_color": "#eef1f5", "theme_color": "#15151c",
                         "id": "/", "scope": "/", "orientation": "portrait",
                         "icons": [{"src": "/icon.png", "sizes": "512x512", "type": "image/png", "purpose": "any"},
                                   {"src": "/icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any"}]},
                        media_type="application/manifest+json")


@app.get("/sw.js")
def service_worker():
    # Minimal worker so Android offers "Install". It deliberately caches nothing:
    # every request still goes to the server, so users never see a stale version.
    return Response("self.addEventListener('install',e=>self.skipWaiting());"
                    "self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));"
                    "self.addEventListener('fetch',()=>{});",
                    media_type="application/javascript", headers={"Cache-Control": "no-cache"})


_ICON192 = None


@app.get("/icon-192.png")
def icon192():
    global _ICON192
    if _ICON192 is None:
        from PIL import Image
        buf = io.BytesIO()
        Image.open(ICON_PATH).convert("RGBA").resize((192, 192), Image.LANCZOS).save(buf, "PNG")
        _ICON192 = buf.getvalue()
    return Response(_ICON192, media_type="image/png")


@app.get("/icon.png")
def icon():
    return FileResponse(ICON_PATH, media_type="image/png")


@app.get("/health")
def health():
    return {"ok": True, "app": ts.APP_NAME, "version": ts.APP_VERSION}


@app.post("/api/signup")
async def signup(request: Request):
    body = await request.json()
    args = body.get("args") or []
    company, name, email, pw = [(str(a) if a is not None else "") for a in (list(args) + ["", "", "", ""])[:4]]
    company, name, email = company.strip(), name.strip(), email.strip()
    if not company or not name or not email or not pw:
        return {"result": {"ok": False, "error": "Company, name, email and password are all needed."}}
    if "@" not in email or "." not in email.split("@")[-1]:
        return {"result": {"ok": False, "error": "Enter a valid email address."}}
    if len(pw) < 8:
        return {"result": {"ok": False, "error": "Use a password of at least 8 characters."}}
    with _acc() as con:
        if con.execute("SELECT 1 FROM directory WHERE email=?", (email.lower(),)).fetchone():
            return {"result": {"ok": False, "error": "That email already has an account. Sign in instead."}}
        cur = con.execute("INSERT INTO companies(name, created_at) VALUES(?,?)",
                          (company, dt.datetime.utcnow().isoformat(timespec="seconds")))
        cid = cur.lastrowid
    os.makedirs(company_dir(cid), exist_ok=True)
    cfg = load_company_cfg(cid)
    cfg["company_name"] = company
    _save_company_cfg(cfg)
    db = company_db(cid)
    db.set_meta("company_name", company)
    db.create_first_admin(name, email, pw)
    sync_directory(cid)
    u = _public_user(db.get_user_by_email(email))
    resp = JSONResponse({"result": {"ok": True, "user": u}})
    set_cookie(resp, make_session(cid, u["id"], 30, u["email"]), 30)
    return resp


@app.post("/api/login")
async def login(request: Request):
    body = await request.json()
    args = body.get("args") or ["", ""]
    ident = str(args[0] if args else "").strip().lower()
    ip = (request.headers.get("x-forwarded-for") or (request.client.host if request.client else "") or "").split(",")[0].strip()
    if _too_many("u:" + ident) or _too_many("ip:" + ip, limit=40):
        return JSONResponse({"error": "Too many sign-in attempts. Wait 15 minutes and try again."}, status_code=429)
    found = await run_in_threadpool(find_login, ident, args[1] if len(args) > 1 else "")
    if not found:
        _record_fail("u:" + ident); _record_fail("ip:" + ip)
        return {"result": None}
    _clear_fail("u:" + ident)
    cid, u = found
    resp = JSONResponse({"result": u})
    set_cookie(resp, make_session(cid, u["id"], 1, u["email"]), 1)   # 'remember' extends it below
    return resp


@app.post("/api/remember_login")
async def remember_login(request: Request):
    cur = current(request)
    if not cur:
        return {"result": {"ok": False}}
    body = await request.json()
    remember = bool((body.get("args") or [False])[0])
    resp = JSONResponse({"result": {"ok": True}})
    set_cookie(resp, make_session(cur[0], cur[1]["id"], 30 if remember else 1, cur[1]["email"]), 30 if remember else 1)
    return resp


@app.post("/api/auto_login")
async def auto_login(request: Request):
    cur = current(request)
    return {"result": cur[1] if cur else None}


@app.post("/api/logout")
@app.post("/api/forget_login")
async def logout(request: Request):
    cur = current(request)
    if cur:
        try:
            company_db(cur[0]).audit(cur[1]["email"], "LOGOUT")
        except Exception:
            pass
    resp = JSONResponse({"result": True})
    clear_cookie(resp)
    return resp


@app.post("/api/brand")
async def brand(request: Request):
    cur = current(request)
    api = WebApi(cur[0], cur[1]) if cur else WebApi()
    return {"result": api.brand()}


# ---------------------------------------------------------------------------
# Import from the desktop app: its SQLite file + a zip of its Receipts folder
# ---------------------------------------------------------------------------
from fastapi import UploadFile, File  # noqa: E402

_ATTACH_COLS = [("expenses", "receipt"), ("expenses", "payment_receipt"), ("job_refs", "quote_attachment"),
                ("job_variations", "attachment"), ("job_invoices", "attachment"),
                ("job_bills", "attachment"), ("job_documents", "attachment")]


def _relink_attachments(cid):
    """Attachment columns hold absolute paths from the old PC (e.g.
    Z:\\Timesheets\\Receipts\\J123\\bill.pdf). Point them at this company's
    Receipts folder, keeping the part after 'Receipts'."""
    rec = os.path.join(company_dir(cid), "Receipts")
    os.makedirs(rec, exist_ok=True)
    rec_norm = rec.replace("\\", "/")
    n = 0
    db = company_db(cid)
    with db.conn() as con:
        for table, col in _ATTACH_COLS:
            try:
                rows = con.execute(f"SELECT id, {col} AS p FROM {table} WHERE {col}<>''").fetchall()
            except Exception:
                continue
            for r in rows:
                p = str(r["p"]).replace("\\", "/")
                if p.startswith(rec_norm):
                    continue
                idx = p.lower().rfind("/receipts/")
                tail = p[idx + len("/receipts/"):] if idx >= 0 else p.rsplit("/", 1)[-1]
                newp = os.path.join(rec, *[t for t in tail.split("/") if t])
                con.execute(f"UPDATE {table} SET {col}=? WHERE id=?", (newp, r["id"]))
                n += 1
    return n


async def _save_upload(up, dest, limit):
    size = 0
    with open(dest, "wb") as f:
        while True:
            chunk = await up.read(1024 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > limit:
                f.close()
                os.remove(dest)
                return None
            f.write(chunk)
    return size


@app.post("/api/import_database")
async def import_database(request: Request, file: UploadFile = File(...)):
    cur = current(request)
    if not cur:
        return JSONResponse({"error": "Not signed in"}, status_code=401)
    cid, user = cur
    if user["user_type"] != "Admin":
        return JSONResponse({"ok": False, "error": "Admins only"}, status_code=403)
    cdir = company_dir(cid)
    tmp = os.path.join(cdir, "import_upload.db")
    if await _save_upload(file, tmp, 512 * 1024 * 1024) is None:
        return {"ok": False, "error": "That file is over 512 MB."}
    try:
        with open(tmp, "rb") as f:
            if f.read(16) != b"SQLite format 3\x00":
                raise ValueError("That is not a SQLite database file. Pick timesheets.db from the desktop app's folder.")
        con = sqlite3.connect(tmp)
        names = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        con.close()
        for need in ("users", "timesheets", "timesheet_lines", "job_refs"):
            if need not in names:
                raise ValueError("That database is missing the '%s' table - it isn't a Timesheets database." % need)
    except Exception as e:
        try:
            os.remove(tmp)
        except Exception:
            pass
        return {"ok": False, "error": str(e)}
    live = os.path.join(cdir, "timesheets.db")
    bk = os.path.join(cdir, "backups")
    os.makedirs(bk, exist_ok=True)
    try:
        if os.path.exists(live):
            shutil.copy2(live, os.path.join(bk, f"pre-import_{dt.datetime.now():%Y%m%d_%H%M%S}.db"))
        with _DBS_LOCK:
            _DBS.pop(cid, None)
        shutil.move(tmp, live)
        db = company_db(cid)                     # init_schema + seed_defaults = migrations
        try:                                      # the file carries its company name (newer desktop builds)
            nm = (db.get_meta("company_name") or "").strip()
            if nm:
                cfg = load_company_cfg(cid); cfg["company_name"] = nm; _save_company_cfg(cfg)
                with _acc() as con:
                    con.execute("UPDATE companies SET name=? WHERE id=?", (nm, cid))
        except Exception:
            pass
        relinked = _relink_attachments(cid)
        sync_directory(cid)
        with db.conn() as con:
            counts = {t: con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                      for t in ("users", "timesheets", "job_refs", "expenses")}
        db.audit(user["email"], "IMPORT_DATABASE", "", f"{file.filename} {counts}")
    except Exception as e:
        return {"ok": False, "error": "Import failed: " + str(e)}
    resp = JSONResponse({"ok": True, "users": counts["users"], "timesheets": counts["timesheets"],
                         "jobs": counts["job_refs"], "expenses": counts["expenses"], "relinked": relinked})
    clear_cookie(resp)        # the importing account no longer exists in this database
    return resp


@app.post("/api/import_receipts")
async def import_receipts(request: Request, file: UploadFile = File(...)):
    import zipfile
    cur = current(request)
    if not cur:
        return JSONResponse({"error": "Not signed in"}, status_code=401)
    cid, user = cur
    if user["user_type"] != "Admin":
        return JSONResponse({"ok": False, "error": "Admins only"}, status_code=403)
    cdir = company_dir(cid)
    tmp = os.path.join(cdir, "import_receipts.zip")
    if await _save_upload(file, tmp, 2 * 1024 * 1024 * 1024) is None:
        return {"ok": False, "error": "That zip is over 2 GB."}
    rec = os.path.abspath(os.path.join(cdir, "Receipts"))
    os.makedirs(rec, exist_ok=True)
    n = 0
    try:
        with zipfile.ZipFile(tmp) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                parts = info.filename.replace("\\", "/").split("/")
                if parts and parts[0].lower() == "receipts":
                    parts = parts[1:]
                parts = [p for p in parts if p not in ("", ".", "..")]
                if not parts:
                    continue
                dest = os.path.abspath(os.path.join(rec, *parts))
                if not dest.startswith(rec + os.sep):
                    continue
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                with z.open(info) as src, open(dest, "wb") as out:
                    shutil.copyfileobj(src, out)
                n += 1
    except zipfile.BadZipFile:
        return {"ok": False, "error": "That isn't a zip file. Right-click the Receipts folder > Send to > Compressed (zipped) folder."}
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass
    relinked = _relink_attachments(cid)
    company_db(cid).audit(user["email"], "IMPORT_RECEIPTS", "", f"{file.filename} ({n} files)")
    return {"ok": True, "files": n, "relinked": relinked}


@app.post("/api/{method}")
async def api_call(method: str, request: Request):
    if method in BLOCKED or method.startswith("_") or not re.fullmatch(r"[a-z][a-z0-9_]*", method):
        raise HTTPException(404)
    cur = current(request)
    if not cur:
        return JSONResponse({"error": "Not signed in"}, status_code=401)
    cid, user = cur
    role = user.get("user_type") or "Standard"
    if method in ADMIN_ONLY and role != "Admin":
        return JSONResponse({"error": "Admins only"}, status_code=403)
    if method in MANAGER_ONLY and role not in ("Manager", "Admin"):
        return JSONResponse({"error": "Managers only"}, status_code=403)
    try:
        body = await request.json()
    except Exception:
        return JSONResponse({"error": "Bad request"}, status_code=400)
    if not isinstance(body, dict):
        return JSONResponse({"error": "Bad request"}, status_code=400)
    args = body.get("args") or []
    if not isinstance(args, list) or len(args) > 20:
        return JSONResponse({"error": "Bad request"}, status_code=400)
    upload = body.get("upload")
    if upload:
        upload = os.path.abspath(str(upload))
        if not upload.startswith(os.path.abspath(company_dir(cid)) + os.sep):
            upload = None
    api = WebApi(cid, user, upload)
    fn = getattr(api, method, None)
    if fn is None or not callable(fn):
        raise HTTPException(404)
    try:
        result = await run_in_threadpool(fn, *args)
    except PermissionError as e:
        return JSONResponse({"error": str(e)}, status_code=403)
    except TypeError as e:
        return JSONResponse({"error": "Bad request: " + str(e)}, status_code=400)
    except Exception as e:
        return JSONResponse({"error": str(e) or e.__class__.__name__}, status_code=500)
    finally:
        if upload and method in ("import_external_timesheet", "import_invoices"):
            try:
                os.remove(upload)
            except Exception:
                pass
    return {"result": result}


_INLINE = {".pdf": "application/pdf", ".png": "image/png", ".jpg": "image/jpeg",
           ".jpeg": "image/jpeg", ".gif": "image/gif"}


@app.get("/api/open")
def open_path(request: Request, p: str = "", dl: str = "", t: str = ""):
    if t:
        d = read_file_token(t)
        if not d:
            raise HTTPException(403, "This link has expired - open the file again from the app.")
        cid, path = d["cid"], os.path.abspath(d["p"])
    else:
        cur = current(request)
        if not cur:
            return RedirectResponse("/")
        cid, _ = cur
        path = os.path.abspath(p or "")
    allowed = (os.path.abspath(company_dir(cid)) + os.sep, os.path.abspath(os.path.join(EXPORTS_DIR, str(cid))) + os.sep)
    if not path.startswith(allowed) or not os.path.isfile(path):
        raise HTTPException(404)
    name = os.path.basename(path)
    if path.startswith(allowed[1]):
        name = name.split("_", 1)[-1] if "_" in name else name      # strip the random prefix
        return FileResponse(path, filename=name, media_type="application/octet-stream")
    ext = os.path.splitext(path)[1].lower()
    if dl:
        return FileResponse(path, filename=name, media_type="application/octet-stream")
    if ext in _INLINE:
        return FileResponse(path, media_type=_INLINE[ext], headers={"Content-Disposition": f'inline; filename="{name}"'})
    return FileResponse(path, filename=name)


# ---------------------------------------------------------------------------
# Housekeeping: daily backups per company, old exports cleared
# ---------------------------------------------------------------------------
def _housekeeping():
    while True:
        try:
            for d in os.listdir(COMPANIES_DIR):
                dbp = os.path.join(COMPANIES_DIR, d, "timesheets.db")
                if os.path.exists(dbp):
                    ts._backup_db(dbp, delay=0)
            cutoff = time.time() - 6 * 3600
            for root, _, files in os.walk(EXPORTS_DIR):
                for f in files:
                    fp = os.path.join(root, f)
                    try:
                        if os.path.getmtime(fp) < cutoff:
                            os.remove(fp)
                    except Exception:
                        pass
        except Exception:
            pass
        time.sleep(6 * 3600)


threading.Thread(target=_housekeeping, daemon=True).start()
