"""
Timesheets desktop client.

A native window onto the hosted Timesheets service: sign-in, data, approvals and
reports all live online. Nothing is stored on this PC except the sign-in cookie.

The server address is built in (SERVER_URL). To point a PC at a different
server (e.g. a test server), put the address in %APPDATA%\\Timesheets\\server.txt
"""
import os, sys, json, datetime as dt, tempfile, shutil, glob, socket, urllib.parse

APP_NAME = "Timesheets"
SERVER_URL = "https://timesheets.kivo.co.uk"      # <- the live service

DATA_DIR = os.path.join(os.environ.get("APPDATA") or tempfile.gettempdir(), APP_NAME)
LOCAL_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), APP_NAME)
for d in (DATA_DIR, LOCAL_DIR):
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass


def server_url():
    try:
        with open(os.path.join(DATA_DIR, "server.txt"), "r", encoding="utf-8") as f:
            u = f.read().strip()
            if u.startswith(("http://", "https://")):
                return u.rstrip("/")
    except Exception:
        pass
    return SERVER_URL


def reachable(url, timeout=4):
    try:
        p = urllib.parse.urlparse(url)
        port = p.port or (443 if p.scheme == "https" else 80)
        with socket.create_connection((p.hostname, port), timeout=timeout):
            return True
    except Exception:
        return False


OFFLINE_HTML = """<body style="font-family:Segoe UI,sans-serif;background:#eef1f5;margin:0;display:flex;align-items:center;justify-content:center;height:100vh">
<div style="background:#fff;border-radius:18px;box-shadow:0 12px 40px rgba(0,0,0,.12);width:420px;overflow:hidden">
<div style="height:6px;background:#84c341"></div><div style="padding:34px 40px">
<div style="font-weight:800;font-size:24px;color:#5e8e1f">Can't reach Timesheets</div>
<p style="color:#555;line-height:1.5;margin-top:10px">This app needs an internet connection. Check you're online, then click Retry.</p>
<p style="color:#999;font-size:12px">Server: %s</p>
<button onclick="location.href='%s'" style="background:#84c341;border:0;color:#15310a;font-weight:700;padding:12px 18px;border-radius:10px;font-size:14px;cursor:pointer">Retry</button>
</div></div></body>"""


def main():
    import webview
    url = server_url()
    target = url + "/?client=1"
    webview.settings["ALLOW_DOWNLOADS"] = True
    webview.settings["OPEN_EXTERNAL_LINKS_IN_BROWSER"] = True
    # Browser profile on local disk so "keep me signed in" survives restarts.
    store = os.path.join(LOCAL_DIR, "profile")
    try:
        os.makedirs(store, exist_ok=True)
    except Exception:
        store = None
    kw = dict(width=1240, height=800, min_size=(1040, 700))
    if reachable(url):
        kw["url"] = target
    else:
        kw["html"] = OFFLINE_HTML % (url, target)
    try:
        webview.create_window(APP_NAME, maximized=True, **kw)
    except TypeError:
        webview.create_window(APP_NAME, **kw)
    icon = os.path.join(os.path.dirname(os.path.abspath(__file__)), "timesheets.ico")
    kws = {"icon": icon} if os.path.exists(icon) else {}
    if store:
        webview.start(private_mode=False, storage_path=store, **kws)
    else:
        webview.start(**kws)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        try:
            with open(os.path.join(DATA_DIR, "client_crash_log.txt"), "a", encoding="utf-8") as f:
                f.write(f"\n===== {dt.datetime.now().isoformat(timespec='seconds')} =====\n")
                f.write(traceback.format_exc())
        except Exception:
            pass
        raise
