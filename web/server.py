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
import os, sys, json, hmac, hashlib, base64, time, re, secrets, shutil, threading
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


def make_session(cid, uid, days):
    exp = int(time.time()) + int(days * 86400)
    return _sign(json.dumps({"cid": int(cid), "uid": int(uid), "exp": exp}).encode())


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
  function openPath(p){ if(!p) return false; const u='/api/open?p='+encodeURIComponent(p);
    if(CLIENT){ let f=document.getElementById('_dlframe'); if(!f){ f=document.createElement('iframe'); f.id='_dlframe'; f.style.display='none'; document.body.appendChild(f); } f.src=u+'&dl=1'; return true; }
    window.open(u,'_blank'); return true; }
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

def build_index_html():
    html = ts.INDEX_HTML.replace("{{LOGO}}", "").replace("{{LOGOCLS}}", "")

    def rep(old, new, count=1):
        nonlocal html
        assert html.count(old) == count, f"page template changed: {old[:80]!r}"
        html = html.replace(old, new)

    # bridge first
    rep("<script>", BRIDGE_JS + "\n<script>")
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
    rep('''async function doSetup(){''', '''function showSignup(){ E('login').classList.add('hidden'); E('setup').classList.remove('hidden'); setTimeout(()=>{const c=E('su_co');if(c)c.focus();},50); }
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
  USER=u; await applyBrand();''')
    # settings: no data-folder card online
    rep('''  <div class="card" style="margin-top:16px"><div style="font-weight:700;margin-bottom:6px">Data</div>''',
        '''  <div class="card" style="margin-top:16px"><div style="font-weight:700;margin-bottom:6px">Your data</div>
    <div class="muted" style="font-size:13px;line-height:1.5">Stored online in your company's own private database, backed up daily.</div>
    <div class="muted" style="font-size:11.5px;margin-top:12px">${esc(window._brand?window._brand.app:'Timesheets')} ${esc(s.version)}</div></div>
  <div class="hidden"><div style="font-weight:700;margin-bottom:6px">Data</div>''')
    rep('''    <div class="k">Open maximised</div>''', '''    <div class="k hidden">Open maximised</div>''')
    rep('''<label style="font-size:13px"><input type="checkbox" id="st_max"''', '''<label class="hidden" style="font-size:13px"><input type="checkbox" id="st_max"''')
    rep('<meta charset="utf-8">',
        '<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Timesheets</title>'
        '<link rel="manifest" href="/manifest.webmanifest"><link rel="icon" href="/icon.png"><link rel="apple-touch-icon" href="/icon.png">'
        '<meta name="theme-color" content="#15151c">')
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
    return s["cid"], _public_user(row)


@app.get("/", response_class=HTMLResponse)
def index():
    return HTMLResponse(INDEX_HTML)


ICON_PATH = os.path.join(os.path.dirname(HERE), "timesheets.png")


@app.get("/manifest.webmanifest")
def manifest():
    return JSONResponse({"name": "Timesheets", "short_name": "Timesheets", "start_url": "/",
                         "display": "standalone", "background_color": "#eef1f5", "theme_color": "#15151c",
                         "icons": [{"src": "/icon.png", "sizes": "512x512", "type": "image/png", "purpose": "any"}]},
                        media_type="application/manifest+json")


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
    db.create_first_admin(name, email, pw)
    sync_directory(cid)
    u = _public_user(db.get_user_by_email(email))
    resp = JSONResponse({"result": {"ok": True, "user": u}})
    set_cookie(resp, make_session(cid, u["id"], 30), 30)
    return resp


@app.post("/api/login")
async def login(request: Request):
    body = await request.json()
    args = body.get("args") or ["", ""]
    found = find_login(args[0] if args else "", args[1] if len(args) > 1 else "")
    if not found:
        return {"result": None}
    cid, u = found
    resp = JSONResponse({"result": u})
    set_cookie(resp, make_session(cid, u["id"], 1), 1)   # 'remember' extends it below
    return resp


@app.post("/api/remember_login")
async def remember_login(request: Request):
    cur = current(request)
    if not cur:
        return {"result": {"ok": False}}
    body = await request.json()
    remember = bool((body.get("args") or [False])[0])
    resp = JSONResponse({"result": {"ok": True}})
    set_cookie(resp, make_session(cur[0], cur[1]["id"], 30 if remember else 1), 30 if remember else 1)
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


@app.post("/api/{method}")
async def api_call(method: str, request: Request):
    if method in BLOCKED or method.startswith("_") or not re.fullmatch(r"[a-z][a-z0-9_]*", method):
        raise HTTPException(404)
    cur = current(request)
    if not cur:
        return JSONResponse({"error": "Not signed in"}, status_code=401)
    cid, user = cur
    body = await request.json()
    args = body.get("args") or []
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
        result = fn(*args)
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
def open_path(request: Request, p: str = "", dl: str = ""):
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
