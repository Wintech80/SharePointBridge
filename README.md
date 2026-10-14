# SharePoint Bridge

Lets tools that can't talk to SharePoint — such as **Beyond Compare** — browse, compare and copy SharePoint Online files.

It runs on your PC as a small local **WebDAV server**. On one side it signs in to Microsoft 365 and talks to SharePoint through the Microsoft Graph API; on the other it looks like an ordinary WebDAV server that Beyond Compare (or any WebDAV client) can connect to.

```
Beyond Compare  <--WebDAV (localhost)-->  SharePoint Bridge  <--Microsoft Graph-->  SharePoint Online
```

## Contents
- [Quick start](#quick-start)
- [Connecting Beyond Compare](#connecting-beyond-compare)
- [The tray icon](#the-tray-icon)
- [Configuration](#configuration)
- [What it does](#what-it-does)
- [Microsoft app registration and security](#microsoft-app-registration-and-security)
- [Running from source](#running-from-source)
- [Building the exe](#building-the-exe)
- [Project layout](#project-layout)
- [Troubleshooting](#troubleshooting)
- [Limitations](#limitations)

## Quick start

1. Double-click **SharePointBridge.exe** (built in `dist\`, see [Building the exe](#building-the-exe)).
2. First run only — a setup window asks:
   - **SharePoint site address** (e.g. `https://contoso.sharepoint.com/sites/Team`), or press Enter to show every site you can access.
   - **Your work email**, only if your tenant can't be worked out from the site address.
   - **Allow writing changes back to SharePoint?** Default is no (read-only).
3. Sign in to Microsoft when shown the code and web page. The bridge registers its own small app in your Microsoft 365 tenant (see [below](#microsoft-app-registration-and-security)).
4. A message box shows the **username and password** Beyond Compare must use. Copy them. (The password can be copied again from the tray icon.)
5. The bridge now runs in the **system tray**. Leave it running while you use Beyond Compare.

Your account must be allowed to register apps in the tenant. If it can't grant consent automatically, the setup prints a link for an administrator to approve; otherwise you approve it yourself at first sign-in.

## Connecting Beyond Compare

Create a WebDAV profile (Folder Compare → click a folder box → Browse Profiles → New Profile):

| Setting | Value |
|---|---|
| Connection type | WebDAV |
| Server / URL | `http://127.0.0.1:8080/` (the port from `config.json`) |
| Username | `bridge` |
| Password | the generated password |

Your SharePoint appears as folders: **site → document library → folders and files**.

Tips
- Beyond Compare appears to transfer one WebDAV file at a time. The bridge's [read-ahead](#what-it-does) works around this.
- Setting a longer connection timeout in the profile helps on slow links.

## The tray icon

Right-click the icon:

| Item | What it does |
|---|---|
| Running at … | Shows the address (informational) |
| Open in browser | Browse your SharePoint through the bridge (asks for the login) |
| View log | Opens `bridge.log` |
| Open data folder | Opens `%USERPROFILE%\.sharepointbrowser` |
| Copy password | Copies the Beyond Compare password to the clipboard |
| Start with Windows | Toggles starting the bridge when you sign in |
| Quit | Stops the bridge |

Closing the bridge is only possible through **Quit** (there is no window). Starting a second copy shows an "already running / port in use" message.

### Command line

The exe also works from a terminal:

```
SharePointBridge.exe              start the bridge (default)
SharePointBridge.exe serve --console   start with a visible console instead of the tray
SharePointBridge.exe setup        re-run setup (change site, read/write, re-register the app)
SharePointBridge.exe login        sign in only
SharePointBridge.exe ls [path]    list SharePoint contents with sizes, dates, authors and SHA1 hashes
```

## Configuration

Settings live in `config.json` in the data folder (`%USERPROFILE%\.sharepointbrowser\config.json` for the exe, the project folder when running from source). Edit it with Notepad and restart the bridge.

| Key | Default | Meaning |
|---|---|---|
| `client_id` | (set by setup) | The Microsoft app registration used to sign in |
| `tenant` | (set by setup) | Your Microsoft 365 tenant id |
| `sites` | `[]` | Site addresses to show. Empty = every site you can access. Bad entries are skipped with a log warning |
| `host` | `127.0.0.1` | Address to listen on. Keep it as localhost |
| `port` | `8080` | Port to listen on |
| `read_only` | `true` | `false` lets Beyond Compare write back to SharePoint. Needs the write permission (re-run `setup` and answer `y`) |
| `require_login` | `true` | Require the username/password below |
| `username` / `password` | `bridge` / generated | Login Beyond Compare must use |
| `download_threads` | `4` | Parallel connections per large file download |
| `readahead` | `3` | Files ahead to fetch in a folder (0 = off) |
| `readahead_max_mb` | `200` | Read-ahead skips files bigger than this |
| `cache_max_gb` | `5` | Download cache size limit |

### Data folder contents

| Item | Purpose |
|---|---|
| `config.json` | Settings (includes the password) |
| `token_cache.json` | Your saved Microsoft sign-in. Delete it to force a fresh login |
| `bridge.log` (+ `.1`–`.3`) | Activity and errors, rotated at 2 MB |
| `downloads\` | Cache of downloaded files |

Setting the environment variable `SPB_HOME` points the data folder somewhere else (useful for testing).

## What it does

**Browsing** — Folder listings are fetched on demand and cached for 60 seconds. Large folders are read page by page. Sites or libraries with the same name get "(2)" added.

**Downloads**
- Streamed: bytes reach Beyond Compare as they arrive from SharePoint, so large files don't time out waiting.
- Files over 16 MB are fetched in 8 MB pieces over several connections at once.
- Interrupted connections resume where they stopped.
- Files are cached (keyed by item and version), so repeat requests are instant and partial-range requests work.
- **Read-ahead:** when a file is requested, the next few files in the same folder are fetched in the background, so a one-file-at-a-time client finds them already cached.

**Writing** (when `read_only` is `false`)
- Upload new files or replace existing ones. Files over 4 MB upload in 10 MB chunks and are streamed to SharePoint while the client is still sending, so progress is real.
- Create folders, delete, rename and move (within one library).
- Copying inside SharePoint downloads then uploads.
- Modified times set by Beyond Compare are passed through to SharePoint.
- Deletes go to the SharePoint recycle bin.
- Names SharePoint won't accept (`" * : < > ? / \ |`, leading/trailing space, trailing dot, `~$`, `_vti_`, reserved names like `CON`, paths over 400 characters) are refused with a clear message.
- SharePoint's own errors are translated: locked/checked-out files, storage quota full, conflicts.

**Housekeeping** — The download cache is kept under `cache_max_gb` by deleting the least recently used files, at startup and then every hour. Abandoned partial downloads and temp upload files are removed too. Files currently in use are never deleted.

## Microsoft app registration and security

Setup creates one app registration named **SharePoint Bridge** in your tenant:
- Single-tenant, public client, redirect URI `http://localhost`.
- **Delegated** Microsoft Graph permissions — it acts only as you, never with its own standing access:
  - `Sites.Read.All`
  - `Files.Read.All` (read-only) **or** `Files.ReadWrite.All` (if you chose write access)
- Admin consent is granted when the signed-in account is allowed to.

To remove it, delete **SharePoint Bridge** under Entra ID → App registrations.

Security notes
- The server listens on `127.0.0.1` only (not reachable from other machines).
- It requires the generated username/password, so other programs or users on your PC can't use your SharePoint access through it.
- `token_cache.json` and `config.json` hold secrets; they are in your user profile. Treat them like a saved login.
- The password is sent over plain HTTP, but only within your own machine.

## Running from source

Requires Python 3.10+.

```
pip install -r requirements.txt
python main.py              # first run does the setup; then serves in the console (Ctrl+C to stop)
python main.py ls "Team Site/Documents"
```

Running from source keeps `config.json` in the project folder and shows no tray icon.

## Building the exe

```
powershell -ExecutionPolicy Bypass -File build_exe.ps1
```

Produces `dist\SharePointBridge.exe`: one file, no console window, with the tray icon. Python is not needed to run it. `build\` and the `.spec` file it creates can be deleted afterwards.

Windows SmartScreen may warn about an unsigned exe the first time: choose **More info → Run anyway**.

## Project layout

| File | Role |
|---|---|
| `main.py` | Entry point: config, setup, commands, server, tray vs console mode |
| `graph.py` | Microsoft Graph access: sign-in, browsing, cached downloads, uploads, read-ahead |
| `provider.py` | Presents SharePoint as a WebDAV file tree (WsgiDAV provider); filename rules and error mapping |
| `register.py` | First-run setup: finds your tenant and registers the app via Graph |
| `housekeeping.py` | Log rotation and download cache clean-up |
| `winutil.py` | Windows helpers: console, message boxes, start-with-Windows |
| `tray.py` | System tray icon and menu |
| `build_exe.ps1` | Builds the exe with PyInstaller |
| `requirements.txt` | Python dependencies |

## Troubleshooting

| Problem | Likely cause / fix |
|---|---|
| Beyond Compare says unauthorized / 401 | Enter the username and password from the setup message (tray → Copy password) |
| `AADSTS700016 … application was not found` | The registered app was deleted or you're in a different tenant. Run `SharePointBridge.exe setup` |
| "port in use / already running" | Another copy is running (check the tray) or another program uses the port; change `port` in `config.json` |
| `502` / "SharePoint rejected the request" on the site list | A bad address in `sites`. Check `bridge.log`; fix or empty the list |
| Writes fail with 403 | `read_only` is `true`, or the app was registered read-only. Run `setup` and answer `y` to writing |
| Downloads time out in Beyond Compare | Raise the profile's connection timeout; check `bridge.log` for `START`/`END` lines |
| Copies slow, one file at a time | Beyond Compare's WebDAV limit; read-ahead helps. Throttling by SharePoint (HTTP 429) is retried automatically |
| Sign-in prompts again | `token_cache.json` was removed or the sign-in expired; sign in again |
| Something else | Open the log from the tray menu, or run `SharePointBridge.exe serve --console` |

## Limitations

- SharePoint Online (Microsoft 365) only; not on-premises SharePoint Server.
- Document libraries only (not lists), and files are handled as plain files — no check-out/check-in or version history browsing.
- Moves can't cross libraries.
- Passwords travel over HTTP on localhost; do not change `host` to expose the bridge on a network.
- Only tested against one tenant; the exe is unsigned.
