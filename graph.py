"""Microsoft Graph access to SharePoint Online: sign-in, tree browsing, downloads."""
import os
import time
import threading
from dataclasses import dataclass, field
from datetime import datetime, timezone
from urllib.parse import urlparse, quote

import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"
READ_SCOPES = ["Sites.Read.All", "Files.Read.All"]
WRITE_SCOPES = ["Sites.Read.All", "Files.ReadWrite.All"]
SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024
CHUNK = 320 * 1024 * 32  # 10 MiB; upload-session chunks must be multiples of 320 KiB
SEGMENT = 8 * 1024 * 1024  # download piece size for parallel range requests
PARALLEL_MIN = 16 * 1024 * 1024  # smaller files are fetched in one piece
DATA_DIR = os.environ.get("SPB_HOME") or os.path.join(os.path.expanduser("~"), ".sharepointbrowser")
CHILD_FIELDS = "id,name,size,folder,file,createdDateTime,lastModifiedDateTime,eTag,cTag,webUrl,parentReference,createdBy,lastModifiedBy"
CACHE_TTL = 60  # seconds that folder listings are reused


def unique_name(name, used):
    """Make a display name safe as a path segment and unique among its siblings."""
    base = (name or "unnamed").replace("/", "-").replace("\\", "-").strip() or "unnamed"
    candidate, i = base, 2
    while candidate.lower() in used:
        candidate = f"{base} ({i})"
        i += 1
    used.add(candidate.lower())
    return candidate


def parse_time(value):
    if not value:
        return None
    return datetime.fromisoformat(value.replace("Z", "+00:00")).timestamp()


@dataclass
class Node:
    kind: str  # root | site | drive | item
    name: str
    is_folder: bool = True
    drive_id: str = ""
    item_id: str = ""
    site_id: str = ""
    parent_id: str = ""
    size: int = 0
    modified: float = None
    created: float = None
    etag: str = ""
    mime: str = ""
    web_url: str = ""
    created_by: str = ""
    modified_by: str = ""
    hashes: dict = field(default_factory=dict)


def item_to_node(it, drive_id):
    f = it.get("file", {})
    return Node(
        kind="item",
        name=it["name"],
        is_folder="folder" in it,
        drive_id=drive_id,
        item_id=it["id"],
        parent_id=it.get("parentReference", {}).get("id", ""),
        size=it.get("size", 0),
        modified=parse_time(it.get("lastModifiedDateTime")),
        created=parse_time(it.get("createdDateTime")),
        etag=(it.get("eTag") or it.get("cTag") or "").strip('"'),
        mime=f.get("mimeType", ""),
        web_url=it.get("webUrl", ""),
        created_by=it.get("createdBy", {}).get("user", {}).get("displayName", ""),
        modified_by=it.get("lastModifiedBy", {}).get("user", {}).get("displayName", ""),
        hashes=f.get("hashes", {}),
    )


class SharePoint:
    def __init__(self, config):
        self.cfg = config
        self.read_only = config.get("read_only", True)
        self.scopes = READ_SCOPES if self.read_only else WRITE_SCOPES
        self._root_ids = {}
        self._jobs = {}
        self._jobs_lock = threading.Lock()
        os.makedirs(DATA_DIR, exist_ok=True)
        self.dl_dir = os.path.join(DATA_DIR, "downloads")
        os.makedirs(self.dl_dir, exist_ok=True)
        self._cache_file = os.path.join(DATA_DIR, "token_cache.json")
        self._tokens = msal.SerializableTokenCache()
        if os.path.exists(self._cache_file):
            self._tokens.deserialize(open(self._cache_file).read())
        self._app = msal.PublicClientApplication(
            config["client_id"],
            authority=f"https://login.microsoftonline.com/{config.get('tenant', 'common')}",
            token_cache=self._tokens,
        )
        self._lock = threading.Lock()
        self._listings = {}
        self._session = requests.Session()

    # ---- auth ----
    def token(self):
        with self._lock:
            accounts = self._app.get_accounts()
            result = None
            if accounts:
                result = self._app.acquire_token_silent(self.scopes, account=accounts[0])
            if not result or "access_token" not in result:
                result = self._app.acquire_token_interactive(self.scopes)
            if "access_token" not in result:
                raise RuntimeError(f"Sign-in failed: {result.get('error_description', result)}")
            if self._tokens.has_state_changed:
                with open(self._cache_file, "w") as fh:
                    fh.write(self._tokens.serialize())
            return result["access_token"]

    def _request(self, url, stream=False, method="GET", headers=None, auth=True, **kw):
        if url.startswith("/"):
            url = GRAPH + url
        kw.setdefault("timeout", (15, 300))  # connect, read-stall
        for attempt in range(5):
            h = dict(headers or {})
            if auth:
                h["Authorization"] = f"Bearer {self.token()}"
            r = self._session.request(method, url, headers=h, stream=stream, **kw)
            if r.status_code in (429, 503, 504):
                time.sleep(int(r.headers.get("Retry-After", 2 ** attempt)))
                continue
            r.raise_for_status()
            return r
        r.raise_for_status()

    def _paged(self, url):
        while url:
            data = self._request(url).json()
            yield from data.get("value", [])
            url = data.get("@odata.nextLink")

    # ---- tree ----
    def root(self):
        return Node(kind="root", name="")

    def children(self, node):
        key = (node.kind, node.site_id, node.drive_id, node.item_id)
        hit = self._listings.get(key)
        if hit and time.time() - hit[0] < CACHE_TTL:
            return hit[1]
        out = list(self._list(node))
        if node.kind == "drive" and out and out[0].parent_id:
            self._root_ids.setdefault(node.drive_id, out[0].parent_id)
        self._listings[key] = (time.time(), out)
        return out

    def _list(self, node):
        used = set()
        if node.kind == "root":
            for url in self.cfg.get("sites") or []:
                u = urlparse(url if "//" in url else "https://" + url)
                path = u.path.rstrip("/")
                try:
                    s = self._request(f"/sites/{u.netloc}:{quote(path)}" if path else f"/sites/{u.netloc}").json()
                except requests.HTTPError as e:
                    print(f"WARNING: skipping site '{url}' - SharePoint rejected it ({e.response.status_code}). "
                          f"Expected something like https://contoso.sharepoint.com/sites/Team")
                    continue
                yield Node("site", unique_name(path.split("/")[-1] or s["displayName"], used), site_id=s["id"], web_url=s.get("webUrl", ""))
            if not self.cfg.get("sites"):
                for s in self._paged("/sites?search=*"):
                    yield Node("site", unique_name(s.get("displayName") or s["name"], used), site_id=s["id"], web_url=s.get("webUrl", ""))
        elif node.kind == "site":
            for d in self._paged(f"/sites/{node.site_id}/drives"):
                yield Node("drive", unique_name(d["name"], used), drive_id=d["id"], web_url=d.get("webUrl", ""),
                           created=parse_time(d.get("createdDateTime")), modified=parse_time(d.get("lastModifiedDateTime")))
        else:
            if node.kind == "drive":
                url = f"/drives/{node.drive_id}/root/children?$top=200&$select={CHILD_FIELDS}"
            else:
                url = f"/drives/{node.drive_id}/items/{node.item_id}/children?$top=200&$select={CHILD_FIELDS}"
            for it in self._paged(url):
                yield item_to_node(it, node.drive_id)

    def resolve(self, parts):
        node = self.root()
        for part in parts:
            for c in self.children(node):
                if c.name == part:
                    node = c
                    break
            else:
                return None
        return node

    # ---- file content ----
    def _cache_path(self, node):
        return os.path.join(self.dl_dir, f"{node.item_id}_{node.etag or node.modified}".replace("/", "_").replace('"', ""))

    def _job(self, node):
        path = self._cache_path(node)
        with self._jobs_lock:
            job = self._jobs.get(path)
            if job and not (job.done and job.error):
                return job
            if os.path.exists(path) and os.path.exists(path + ".done"):
                try:
                    os.utime(path + ".done")  # mark as recently used for cache clean-up
                except OSError:
                    pass
                job = DownloadJob(self, node, path, complete=True)
            else:
                job = DownloadJob(self, node, path)
            self._jobs[path] = job
            return job

    def active_paths(self):
        """Cache files that are being downloaded or read right now (must not be cleaned up)."""
        with self._jobs_lock:
            return {p for p, j in self._jobs.items() if not j.done or time.time() - j.last_used < 600}

    def readahead(self, node):
        """Start fetching the next few files in the same folder, so a one-file-at-a-time client finds them already cached."""
        count = int(self.cfg.get("readahead", 3))
        limit = int(self.cfg.get("readahead_max_mb", 200)) * 1048576
        siblings = self._listing_for(node.drive_id, node.parent_id) if count > 0 else None
        if not siblings:
            return
        files = sorted((c for c in siblings if not c.is_folder), key=lambda c: c.name.lower())
        try:
            i = next(k for k, c in enumerate(files) if c.item_id == node.item_id)
        except StopIteration:
            return
        with self._jobs_lock:
            running = sum(1 for j in self._jobs.values() if j.prefetch and not j.done)
        for c in files[i + 1:i + 1 + count]:
            if running >= count:
                break
            if c.size == 0 or c.size > limit:
                continue
            path = self._cache_path(c)
            if path in self._jobs or os.path.exists(path + ".done"):
                continue
            self._job(c).prefetch = True
            running += 1

    def open_download(self, node):
        """File-like object that serves the file while it is still being fetched from SharePoint."""
        job = self._job(node)
        job.last_used = time.time()
        job.wait_started()
        return DownloadReader(job)

    def local_copy(self, node):
        """Download a file once (per version) and return the cached path (blocks until complete)."""
        job = self._job(node)
        job.wait_done()
        return job.path

    # ---- writing ----
    def _cache_edit(self, drive_id, folder_id, fn):
        """Apply fn(list)->list to the cached listing of a folder, if we have one (no re-fetch from SharePoint)."""
        for key, (ts, lst) in list(self._listings.items()):
            if self._is_listing_of(key, drive_id, folder_id):
                self._listings[key] = (ts, fn(lst))

    def _is_listing_of(self, key, drive_id, folder_id):
        kind, _site, drive, item = key
        return drive == drive_id and (item == folder_id or (kind == "drive" and folder_id == self._root_ids.get(drive_id)))

    def _listing_for(self, drive_id, folder_id):
        for key, (_ts, lst) in self._listings.items():
            if self._is_listing_of(key, drive_id, folder_id):
                return lst
        return None

    def _cache_put(self, drive_id, folder_id, node):
        self._cache_edit(drive_id, folder_id, lambda lst: [n for n in lst if n.name != node.name] + [node])

    def _cache_remove(self, drive_id, folder_id, item_id):
        self._cache_edit(drive_id, folder_id, lambda lst: [n for n in lst if n.item_id != item_id])

    def real_id(self, node):
        """Graph item id of a folder node (a drive's root has no id in our tree until asked for)."""
        if node.kind == "item":
            return node.item_id
        if node.kind != "drive":
            raise PermissionError("Can only write inside a document library")
        if node.drive_id not in self._root_ids:
            self._root_ids[node.drive_id] = self._request(f"/drives/{node.drive_id}/root").json()["id"]
        return self._root_ids[node.drive_id]

    def upload(self, parent, name, fh, size):
        """Upload (create or replace) a file named `name` inside folder node `parent`. Returns the new Node."""
        base = f"/drives/{parent.drive_id}/items/{self.real_id(parent)}:/{quote(name)}:"
        if size <= SIMPLE_UPLOAD_MAX:
            fh.seek(0)
            r = self._request(base + "/content?@microsoft.graph.conflictBehavior=replace", method="PUT",
                              data=fh.read(), headers={"Content-Type": "application/octet-stream"})
            item = r.json()
        else:
            fh.seek(0)
            up = self.upload_stream(parent, name, size)
            for chunk in iter(lambda: fh.read(CHUNK), b""):
                up.write(chunk)
            up.close()
            item = up.item
        node = item_to_node(item, parent.drive_id)
        self._cache_put(parent.drive_id, self.real_id(parent), node)
        return node

    def upload_stream(self, parent, name, total):
        """File-like object that forwards data to SharePoint as it is written (total size must be known)."""
        return UploadStream(self, parent, name, total)

    def mkdir(self, parent, name):
        item = self._request(f"/drives/{parent.drive_id}/items/{self.real_id(parent)}/children", method="POST",
                             json={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}).json()
        self._cache_put(parent.drive_id, self.real_id(parent), item_to_node(item, parent.drive_id))

    def delete(self, node):
        if node.kind != "item":
            raise PermissionError("Only files and folders can be deleted")
        self._request(f"/drives/{node.drive_id}/items/{node.item_id}", method="DELETE")
        self._cache_remove(node.drive_id, node.parent_id, node.item_id)

    def move(self, node, new_parent, new_name):
        if node.kind != "item" or new_parent.drive_id != node.drive_id:
            raise PermissionError("Can only move files/folders within the same document library")
        new_parent_id = self.real_id(new_parent)
        item = self._request(f"/drives/{node.drive_id}/items/{node.item_id}", method="PATCH",
                             json={"name": new_name, "parentReference": {"id": new_parent_id}}).json()
        self._cache_remove(node.drive_id, node.parent_id, node.item_id)
        self._cache_put(node.drive_id, new_parent_id, item_to_node(item, node.drive_id))

    def set_modified(self, node, ts):
        if node.kind != "item" or not node.item_id:
            return
        iso = datetime.fromtimestamp(ts, timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        item = self._request(f"/drives/{node.drive_id}/items/{node.item_id}", method="PATCH",
                             json={"fileSystemInfo": {"lastModifiedDateTime": iso}}).json()
        self._cache_put(node.drive_id, node.parent_id, item_to_node(item, node.drive_id))


class UploadStream:
    """Writes a large file to SharePoint through an upload session while the data is still arriving."""

    def __init__(self, sp, parent, name, total):
        self.sp, self.parent, self.name, self.total = sp, parent, name, total
        self.buf = bytearray()
        self.pos = 0
        self.item = None
        self.url = None
        self.closed = False

    def _start(self):
        base = f"/drives/{self.parent.drive_id}/items/{self.sp.real_id(self.parent)}:/{quote(self.name)}:"
        sess = self.sp._request(base + "/createUploadSession", method="POST",
                                json={"item": {"@microsoft.graph.conflictBehavior": "replace"}}).json()
        self.url = sess["uploadUrl"]

    def _send(self, data):
        if self.url is None:
            self._start()
        end = self.pos + len(data) - 1
        r = self.sp._request(self.url, method="PUT", data=bytes(data), auth=False,
                             headers={"Content-Range": f"bytes {self.pos}-{end}/{self.total}"})
        self.pos = end + 1
        if r.status_code in (200, 201):
            self.item = r.json()
        print(f"  uploading {self.name}: {self.pos * 100 // self.total}% ({self.pos / 1048576:.1f} / {self.total / 1048576:.1f} MB)", flush=True)

    def write(self, data):
        self.buf += data
        while len(self.buf) >= CHUNK and self.pos + CHUNK < self.total:
            self._send(self.buf[:CHUNK])
            del self.buf[:CHUNK]
        return len(data)

    def close(self):
        if self.closed:
            return
        self.closed = True
        if self.pos + len(self.buf) != self.total:
            self.abort()
            raise IOError(f"Received {self.pos + len(self.buf)} bytes but expected {self.total}")
        self._send(self.buf)
        self.buf = bytearray()
        self.sp._cache_put(self.parent.drive_id, self.sp.real_id(self.parent), item_to_node(self.item, self.parent.drive_id))

    def abort(self):
        self.closed = True
        if self.url:
            try:
                self.sp._request(self.url, method="DELETE", auth=False)
            except Exception:
                pass


class DownloadJob:
    """Downloads one file from SharePoint into the cache using several parallel range requests,
    exposing the contiguous prefix (`downloaded`) so it can be streamed while the rest is still arriving."""

    def __init__(self, sp, node, path, complete=False):
        self.sp, self.node, self.path = sp, node, path
        self.total = node.size
        self.downloaded = self.total if complete else 0
        self.done = complete
        self.error = None
        self.prefetch = False
        self.last_used = time.time()
        self.cond = threading.Condition()
        self.segs = []
        self.seg_done = []
        self._next = 0
        self._next_lock = threading.Lock()
        if not complete:
            if os.path.exists(path + ".done"):
                os.remove(path + ".done")
            threading.Thread(target=self._run, daemon=True).start()

    def _watermark(self):
        w = 0
        for (s, e), d in zip(self.segs, self.seg_done):
            w += d
            if d < e - s + 1:
                break
        return w

    def _run(self):
        n = self.node
        try:
            if self.total == 0:
                open(self.path, "wb").close()
            else:
                if self.total < PARALLEL_MIN:
                    self.segs = [(0, self.total - 1)]
                else:
                    self.segs = [(s, min(s + SEGMENT, self.total) - 1) for s in range(0, self.total, SEGMENT)]
                self.seg_done = [0] * len(self.segs)
                workers = max(1, min(int(self.sp.cfg.get("download_threads", 4)), len(self.segs)))
                self._big = self.total > 20 * 1024 * 1024
                self._step = max(self.total // 10, 1)
                self._report = self._step
                if self._big:
                    print(f"  downloading {n.name} ({self.total / 1048576:.1f} MB, {workers} connections)...", flush=True)
                with open(self.path, "wb") as fh:
                    fh.truncate(self.total)
                threads = [threading.Thread(target=self._worker, daemon=True) for _ in range(workers)]
                for t in threads:
                    t.start()
                for t in threads:
                    t.join()
                if self.error is None and self.downloaded != self.total:
                    raise IOError(f"Downloaded {self.downloaded} bytes but SharePoint reported {self.total}")
            if self.error is None:
                open(self.path + ".done", "w").close()
        except Exception as e:
            self.error = self.error or e
        finally:
            with self.cond:
                self.done = True
                self.cond.notify_all()

    def _worker(self):
        try:
            while self.error is None:
                with self._next_lock:
                    idx = self._next
                    self._next += 1
                if idx >= len(self.segs):
                    return
                self._fetch(idx)
        except Exception as e:
            self.error = self.error or e
            with self.cond:
                self.cond.notify_all()

    def _fetch(self, idx):
        start, end = self.segs[idx]
        n = self.node
        retries = 0
        with open(self.path, "r+b") as fh:
            while self.seg_done[idx] < end - start + 1 and self.error is None:
                pos = start + self.seg_done[idx]
                try:
                    r = self.sp._request(f"/drives/{n.drive_id}/items/{n.item_id}/content", stream=True,
                                         headers={"Range": f"bytes={pos}-{end}"})
                    if r.status_code != 206 and pos > 0:
                        raise IOError("SharePoint ignored the Range request")
                    fh.seek(pos)
                    for chunk in r.iter_content(256 * 1024):
                        chunk = chunk[: end - pos + 1]  # never write past this segment
                        fh.write(chunk)
                        fh.flush()
                        pos += len(chunk)
                        with self.cond:
                            self.seg_done[idx] += len(chunk)
                            self.downloaded = self._watermark()
                            self.cond.notify_all()
                            got = sum(self.seg_done)
                        if self._big and got >= self._report:
                            print(f"  downloading {n.name}: {got * 100 // self.total}%", flush=True)
                            self._report += self._step
                        if pos > end:
                            break
                except (requests.ConnectionError, requests.Timeout, requests.exceptions.ChunkedEncodingError):
                    retries += 1
                    if retries > 5:
                        raise
                    time.sleep(2)  # resume this segment from where it got to

    def wait_started(self):
        with self.cond:
            self.cond.wait_for(lambda: self.downloaded > 0 or self.done, timeout=120)
        if self.error:
            raise self.error

    def wait_done(self):
        with self.cond:
            self.cond.wait_for(lambda: self.done)
        if self.error:
            raise self.error


class DownloadReader:
    """Seekable, blocking reader over a file that may still be downloading."""

    def __init__(self, job):
        self.job = job
        self.f = open(job.path, "rb")
        self.pos = 0

    def read(self, n=-1):
        job = self.job
        end = job.total if n is None or n < 0 else min(job.total, self.pos + n)
        if end <= self.pos:
            return b""
        with job.cond:
            ok = job.cond.wait_for(lambda: job.downloaded >= end or job.done, timeout=300)
        if not ok:
            raise IOError("Timed out waiting for SharePoint download")
        if job.downloaded < end:
            if job.error:
                raise IOError(f"Download from SharePoint failed: {job.error}")
            end = job.downloaded
        self.f.seek(self.pos)
        data = self.f.read(max(end - self.pos, 0))
        self.pos += len(data)
        return data

    def seek(self, offset, whence=0):
        self.pos = offset if whence == 0 else self.pos + offset if whence == 1 else self.job.total + offset
        return self.pos

    def tell(self):
        return self.pos

    def close(self):
        self.f.close()
