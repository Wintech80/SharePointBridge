"""WebDAV provider that presents SharePoint (via graph.SharePoint) as a file tree.
Writable unless config has "read_only": true."""
import mimetypes
import os
import re
import posixpath
import tempfile
from email.utils import parsedate_to_datetime
from datetime import datetime

from requests import HTTPError
from wsgidav.dav_error import (DAVError, HTTP_FORBIDDEN, HTTP_BAD_GATEWAY, HTTP_NOT_FOUND, HTTP_CONFLICT, HTTP_BAD_REQUEST,
                               HTTP_LOCKED, HTTP_INSUFFICIENT_STORAGE, HTTP_REQUEST_ENTITY_TOO_LARGE, HTTP_PRECONDITION_FAILED)
from wsgidav.dav_provider import DAVCollection, DAVNonCollection, DAVProvider
from wsgidav.util import join_uri

from graph import Node, SIMPLE_UPLOAD_MAX, item_to_node


BAD_CHARS = re.compile(r'["*:<>?/\\|]')
RESERVED = {".lock", "con", "prn", "aux", "nul", "desktop.ini"} | {f"com{i}" for i in range(10)} | {f"lpt{i}" for i in range(10)}
MAX_PATH = 400


def name_problem(name, full_path=""):
    """Why SharePoint would refuse this file/folder name, or None if it is fine."""
    if BAD_CHARS.search(name):
        return f'"{name}" contains a character SharePoint does not allow (one of " * : < > ? / \\ |)'
    if name != name.strip():
        return f'"{name}" starts or ends with a space, which SharePoint does not allow'
    if name.endswith("."):
        return f'"{name}" ends with a dot, which SharePoint does not allow'
    if name.startswith("~$"):
        return f'"{name}" starts with "~$", which SharePoint does not allow'
    if "_vti_" in name.lower():
        return f'"{name}" contains "_vti_", which SharePoint reserves'
    if name.lower() in RESERVED or name.lower().split(".")[0] in RESERVED - {".lock", "desktop.ini"}:
        return f'"{name}" is a reserved name in SharePoint'
    if len(full_path) > MAX_PATH:
        return f"The path is {len(full_path)} characters long; SharePoint allows at most {MAX_PATH}"
    return None


def check_name(name, parent_path=""):
    problem = name_problem(name, posixpath.join(parent_path, name))
    if problem:
        raise DAVError(HTTP_BAD_REQUEST, problem)


STATUS_MAP = {
    400: (HTTP_BAD_REQUEST, "SharePoint rejected the request"),
    403: (HTTP_FORBIDDEN, "SharePoint denied access"),
    404: (HTTP_NOT_FOUND, "Not found in SharePoint"),
    409: (HTTP_CONFLICT, "Conflict in SharePoint (name already exists or item changed)"),
    412: (HTTP_PRECONDITION_FAILED, "The file changed in SharePoint"),
    413: (HTTP_REQUEST_ENTITY_TOO_LARGE, "File too large for SharePoint"),
    423: (HTTP_LOCKED, "The file is locked or checked out by someone else in SharePoint"),
    507: (HTTP_INSUFFICIENT_STORAGE, "SharePoint storage quota exceeded"),
}


def guard(fn):
    """Translate Graph/permission errors into WebDAV errors."""
    def wrapper(*a, **kw):
        try:
            return fn(*a, **kw)
        except PermissionError as e:
            raise DAVError(HTTP_FORBIDDEN, str(e))
        except HTTPError as e:
            code = e.response.status_code if e.response is not None else 502
            detail = ""
            try:
                detail = e.response.json()["error"]["message"]
            except Exception:
                pass
            if code in STATUS_MAP:
                status, default = STATUS_MAP[code]
                raise DAVError(status, f"{default}: {detail}" if detail else default)
            raise DAVError(HTTP_BAD_GATEWAY, f"SharePoint error {code}: {detail}")
    return wrapper


def parse_dav_time(text):
    try:
        return parsedate_to_datetime(text).timestamp()
    except (TypeError, ValueError):
        return datetime.fromisoformat(text.replace("Z", "+00:00")).timestamp()


class _Common:
    """Behaviour shared by files and folders."""

    def get_display_name(self):
        return self.node.name or "SharePoint"

    def get_creation_date(self):
        return self.node.created

    def get_last_modified(self):
        return self.node.modified

    def _parent_and_name(self, dest_path):
        parts = [p for p in dest_path.split("/") if p]
        parent = self.sp.resolve(parts[:-1])
        if parent is None:
            raise DAVError(HTTP_CONFLICT, "Destination folder does not exist")
        check_name(parts[-1], "/".join(parts[:-1]))
        return parent, parts[-1]

    @guard
    def handle_move(self, dest_path):
        if self.sp.read_only:
            raise DAVError(HTTP_FORBIDDEN, "Read-only")
        parent, name = self._parent_and_name(dest_path)
        self.sp.move(self.node, parent, name)
        return True

    @guard
    def delete(self):
        self.sp.delete(self.node)

    @guard
    def set_property_value(self, name, value, *, dry_run=False):
        # Beyond Compare sets the modified time after uploading; accept that and ignore other properties.
        if value is not None and name.endswith(("}getlastmodified", "}Win32LastModifiedTime")) and not dry_run:
            self.sp.set_modified(self.node, parse_dav_time(value.text))
        return True

    @guard
    def set_last_modified(self, dest_path, time_stamp, *, dry_run):
        if not dry_run:
            self.sp.set_modified(self.node, parse_dav_time(time_stamp))


class SPCollection(_Common, DAVCollection):
    def __init__(self, path, environ, sp, node):
        DAVCollection.__init__(self, path, environ)
        self.sp, self.node = sp, node

    @guard
    def get_member_names(self):
        return [c.name for c in self.sp.children(self.node)]

    @guard
    def get_member(self, name):
        for c in self.sp.children(self.node):
            if c.name == name:
                return make_resource(join_uri(self.path, name), self.environ, self.sp, c)
        return None

    def support_recursive_delete(self):
        return True

    def support_recursive_move(self, dest_path):
        return True

    @guard
    def create_empty_resource(self, name):
        """Placeholder for a file about to be PUT; nothing is created in SharePoint until the upload finishes."""
        check_name(name, self.path)
        node = Node("item", name, is_folder=False, drive_id=self.node.drive_id)
        return SPFile(join_uri(self.path, name), self.environ, self.sp, node, parent=self.node)

    @guard
    def create_collection(self, name):
        check_name(name, self.path)
        self.sp.mkdir(self.node, name)

    @guard
    def copy_move_single(self, dest_path, *, is_move):
        parent, name = self._parent_and_name(dest_path)
        if not any(c.name == name for c in self.sp.children(parent)):
            self.sp.mkdir(parent, name)


class SPFile(_Common, DAVNonCollection):
    def __init__(self, path, environ, sp, node, parent=None):
        DAVNonCollection.__init__(self, path, environ)
        self.sp, self.node, self._parent = sp, node, parent
        self._tmp = None
        self._stream = None

    def get_content_length(self):
        return self.node.size

    def get_content_type(self):
        return self.node.mime or mimetypes.guess_type(self.node.name)[0] or "application/octet-stream"

    def get_etag(self):
        return self.node.etag or None

    def support_etag(self):
        return bool(self.node.etag)

    def support_ranges(self):
        return True

    @guard
    def get_content(self):
        stream = self.sp.open_download(self.node)
        try:
            self.sp.readahead(self.node)
        except Exception:
            pass  # read-ahead is only an optimisation
        return stream

    def _parent_node(self):
        if self._parent is None:
            parts = [p for p in self.path.split("/") if p]
            self._parent = self.sp.resolve(parts[:-1])
        return self._parent

    def begin_write(self, *, content_type=None):
        if self.sp.read_only:
            raise DAVError(HTTP_FORBIDDEN, "Read-only")
        total = int(self.environ.get("CONTENT_LENGTH") or 0)
        if total > SIMPLE_UPLOAD_MAX:
            # Large file with a known size: forward to SharePoint while Beyond Compare is still sending,
            # so its progress bar reflects the real upload.
            self._stream = self.sp.upload_stream(self._parent_node(), self.node.name, total)
            return self._stream
        # Small/unknown size: WsgiDAV closes the stream before end_write, so keep it on disk and reopen it by name.
        fd = tempfile.NamedTemporaryFile(delete=False, dir=self.sp.dl_dir, suffix=".upload")
        self._tmp = fd.name
        return fd

    @guard
    def end_write(self, *, with_errors):
        stream, self._stream = self._stream, None
        if stream:
            if with_errors:
                stream.abort()
            else:
                self.node = item_to_node(stream.item, self.node.drive_id)
            return
        path, self._tmp = self._tmp, None
        try:
            if not with_errors:
                with open(path, "rb") as fh:
                    self.node = self.sp.upload(self._parent_node(), self.node.name, fh, os.path.getsize(path))
        finally:
            os.remove(path)

    @guard
    def copy_move_single(self, dest_path, *, is_move):
        parent, name = self._parent_and_name(dest_path)
        with open(self.sp.local_copy(self.node), "rb") as fh:
            self.sp.upload(parent, name, fh, self.node.size)


def make_resource(path, environ, sp, node):
    if node.is_folder:
        return SPCollection(path, environ, sp, node)
    return SPFile(path, environ, sp, node)


class SharePointProvider(DAVProvider):
    def __init__(self, sp):
        super().__init__()
        self.sp = sp

    def is_readonly(self):
        return self.sp.read_only

    @guard
    def get_resource_inst(self, path, environ):
        parts = [p for p in path.split("/") if p]
        node = self.sp.resolve(parts)
        if node is None:
            return None
        return make_resource(path or "/", environ, self.sp, node)
