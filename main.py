"""SharePoint -> WebDAV bridge so tools like Beyond Compare can browse/compare SharePoint files.

    SharePointBridge serve     start the local WebDAV server (default; runs first-time setup if needed)
    SharePointBridge setup     (re)run setup: register the Microsoft app and write a fresh config
    SharePointBridge login     sign in only
    SharePointBridge ls [path] list SharePoint contents with full details
    --console                  with serve: run in a console window instead of the system tray
"""
import json
import os
import secrets
import sys
import threading
from datetime import datetime

from cheroot import wsgi
from wsgidav.wsgidav_app import WsgiDAVApp

import housekeeping
import winutil
from graph import SharePoint, DATA_DIR
from register import register, find_tenant
from provider import SharePointProvider

FROZEN = getattr(sys, "frozen", False)
# Packaged exe keeps its settings in the per-user data folder; running from source uses the project folder.
CONFIG = os.path.join(DATA_DIR if FROZEN else os.path.dirname(os.path.abspath(__file__)), "config.json")
DEFAULTS = {
    "client_id": "",  # filled in by first-run setup
    "tenant": "common",
    "sites": [],  # e.g. ["https://contoso.sharepoint.com/sites/Team"]; empty = every site you can see
    "host": "127.0.0.1",
    "port": 8080,
    "read_only": True,  # set to false to allow Beyond Compare to write changes back to SharePoint
    "require_login": True,  # Beyond Compare must send the username/password below
    "username": "bridge",
    "password": "",  # generated automatically
    "download_threads": 4,  # parallel connections per large file download
    "readahead": 3,  # fetch this many following files in the folder ahead of time (0 = off)
    "readahead_max_mb": 200,  # ...but never files bigger than this
    "cache_max_gb": 5,  # download cache size limit; least recently used files are deleted
}


def load_config():
    if not os.path.exists(CONFIG):
        return dict(DEFAULTS)
    with open(CONFIG, encoding="utf-8-sig") as fh:
        return {**DEFAULTS, **json.load(fh)}


def save_config(cfg):
    os.makedirs(os.path.dirname(CONFIG), exist_ok=True)
    with open(CONFIG, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2)


def ensure_password(cfg):
    """Make sure a login password exists. Returns (cfg, newly_generated)."""
    if cfg["require_login"] and not cfg["password"]:
        cfg = {**cfg, "password": secrets.token_urlsafe(12)}
        save_config(cfg)
        return cfg, True
    return cfg, False


def credentials_text(cfg):
    return (f"Beyond Compare login for this bridge:\n    Username: {cfg['username']}\n    Password: {cfg['password']}\n"
            "(Enter these in your WebDAV profile in Beyond Compare.)")


def ask(prompt, default=""):
    try:
        return input(prompt).strip() or default
    except EOFError:
        return default


def run_setup(cfg):
    """Interactive first-run setup: ask a few questions, register the Microsoft app, write config."""
    winutil.ensure_console()
    print("=== SharePoint Bridge setup ===")
    print("This signs you in to Microsoft 365 and registers a small app in your tenant so the bridge can read SharePoint.\n")
    site = ask("SharePoint site address (e.g. https://contoso.sharepoint.com/sites/Team)\nor just press Enter to show every site you can access: ")
    if site and ".sharepoint.com" not in site.lower():
        print("That doesn't look like a SharePoint site address, so it was ignored - all sites will be shown.")
        print("(You can set one later in config.json.)\n")
        site = ""
    email = ""
    tenant = find_tenant(site)
    if not tenant:
        email = ask("Your Microsoft 365 work email (only used to find your tenant): ")
        tenant = find_tenant(email)
    if not tenant:
        tenant = ask("Could not detect your tenant. Enter your tenant ID or domain (e.g. contoso.onmicrosoft.com): ")
    writable = ask("Allow writing changes back to SharePoint? [y/N]: ").lower().startswith("y")
    client_id, tenant, _ = register(tenant, read_only=not writable)
    cfg = {**cfg, "client_id": client_id, "tenant": tenant, "sites": [site] if site else [], "read_only": not writable}
    cfg, _ = ensure_password(cfg)
    save_config(cfg)
    print(f"\nSetup complete. Settings saved to {CONFIG}\n")
    print(credentials_text(cfg) + "\n")
    return cfg


def fmt(ts):
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S") if ts else "-"


def cmd_ls(sp, arg):
    parts = [p for p in arg.split("/") if p]
    node = sp.resolve(parts)
    if node is None:
        sys.exit(f"Not found: {arg}")
    for c in sp.children(node):
        kind = "<DIR>" if c.is_folder else f"{c.size:>12,}"
        print(f"{kind:>12}  {fmt(c.modified)}  {c.name}")
        if not c.is_folder:
            print(f"{'':14}created {fmt(c.created)} by {c.created_by or '-'}; modified by {c.modified_by or '-'}")
            print(f"{'':14}sha1 {c.hashes.get('sha1Hash', '-')}  {c.web_url}")


def build_server(sp, cfg):
    log = housekeeping.log
    if cfg["require_login"]:
        users = {"*": {cfg["username"]: {"password": cfg["password"]}}}
    else:
        users = {"*": True}
    dav = WsgiDAVApp({
        "host": cfg["host"],
        "port": cfg["port"],
        "provider_mapping": {"/": SharePointProvider(sp)},
        "simple_dc": {"user_mapping": users},
        "http_authenticator": {"domain_controller": None, "accept_basic": True, "accept_digest": False,
                               "default_to_digest": False},
        "verbose": 1,
        "dir_browser": {"enable": True},
        "lock_storage": False,
    })

    def app(environ, start_response):
        method, path = environ["REQUEST_METHOD"], environ["PATH_INFO"]
        if method in ("GET", "HEAD"):
            environ.pop("HTTP_DEPTH", None)  # Beyond Compare sends Depth: 1 on GET; WsgiDAV rejects it
        transfer = method in ("GET", "PUT", "POST") and not path.endswith("/")
        thread = threading.current_thread().name
        if transfer:
            log.info(f"START {method} {path} thread={thread} length={environ.get('CONTENT_LENGTH')} range={environ.get('HTTP_RANGE')}")

        def sr(status, headers, exc=None):
            if not status.startswith(("2", "3")) and not (status.startswith("401") and "HTTP_AUTHORIZATION" not in environ):
                log.info(f"{status} {method} {path}")
            return start_response(status, headers, exc)

        try:
            result = dav(environ, sr)
        except BaseException as e:
            log.info(f"EXCEPTION {method} {path}: {e!r}")
            raise
        if not transfer:
            return result

        def finish():
            try:
                yield from result
            finally:
                if hasattr(result, "close"):
                    result.close()
                log.info(f"END   {method} {path} thread={thread}")
        return finish()

    return wsgi.Server((cfg["host"], cfg["port"]), app, timeout=600, numthreads=32)


def cmd_serve(sp, cfg, tray_mode):
    log = housekeeping.log
    housekeeping.start_cache_cleaner(sp, cfg["cache_max_gb"])
    server = build_server(sp, cfg)
    url = f"http://{cfg['host']}:{cfg['port']}/"
    try:
        server.prepare()  # binds the port now, so "already running" errors show up immediately
    except OSError as e:
        raise RuntimeError(f"Could not start on port {cfg['port']} - is the bridge already running (check the system tray)? ({e})")
    log.info(f"SharePoint bridge running at {url}" + ("" if not cfg["read_only"] else " (read-only)"))

    if tray_mode:
        import tray
        threading.Thread(target=server.serve, daemon=True, name="server").start()
        tray.run_tray(url, DATA_DIR, os.path.join(DATA_DIR, "bridge.log"), cfg["password"] if cfg["require_login"] else "", server.stop)
    else:
        print(f"SharePoint bridge running at {url}  (Ctrl+C to stop)")
        try:
            server.serve()
        except KeyboardInterrupt:
            server.stop()


def main():
    args = [a for a in sys.argv[1:] if a != "--console"]
    console_flag = len(args) != len(sys.argv) - 1
    cmd = args[0] if args else "serve"
    if cmd not in ("serve", "setup", "login", "ls"):
        winutil.ensure_console()
        sys.exit(__doc__)
    if cmd != "serve" or console_flag:
        winutil.ensure_console()  # windowed exe: attach to the terminal / open a window for prompts and output
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(line_buffering=True)  # show prompts/sign-in codes immediately, even when piped
        except Exception:
            pass
    no_console = not winutil.has_console()  # windowed exe running in the background (tray mode)
    housekeeping.setup_logging(DATA_DIR, to_console=not no_console)
    if no_console:
        sys.stdout = sys.stderr = housekeeping.LogWriter()  # no console: print() output goes to the log

    cfg = load_config()
    if cmd == "setup" or not cfg["client_id"]:
        cfg = run_setup(cfg)
        if cmd == "setup":
            return
        if no_console:  # first double-click run: setup used a temporary console window; carry on in the tray
            ask("Press Enter to start the bridge (it keeps running in the system tray)...")
            winutil.release_console()
            sys.stdout = sys.stderr = housekeeping.LogWriter()
    cfg, generated = ensure_password(cfg)
    if generated:
        text = credentials_text(cfg)
        if not no_console:
            print(text + "\n")
        else:
            winutil.message_box("SharePoint Bridge - Beyond Compare login", text + "\n\nYou can copy the password later from the tray icon menu.")
    sp = SharePoint(cfg)
    sp.token()  # sign in up front (opens a browser the first time)
    if cmd == "ls":
        cmd_ls(sp, args[1] if len(args) > 1 else "")
    elif cmd == "serve":
        cmd_serve(sp, cfg, tray_mode=no_console)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    except SystemExit:
        raise
    except Exception as e:
        try:
            housekeeping.log.info(f"ERROR: {e!r}")
        except Exception:
            pass
        if winutil.has_console():
            print(f"\nERROR: {e}")
            if FROZEN:
                ask("Press Enter to close...")
        else:
            winutil.message_box("SharePoint Bridge", f"Error: {e}", error=True)
        sys.exit(1)
