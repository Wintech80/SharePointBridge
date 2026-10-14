"""Logging with rotation, and keeping the download cache under a size limit."""
import logging
import logging.handlers
import os
import sys
import threading
import time

log = logging.getLogger("bridge")


def setup_logging(data_dir, to_console):
    os.makedirs(data_dir, exist_ok=True)
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter("%(asctime)s %(message)s", "%H:%M:%S")
    fh = logging.handlers.RotatingFileHandler(os.path.join(data_dir, "bridge.log"), maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    if to_console and sys.stdout is not None:
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(fmt)
        log.addHandler(ch)
    for stale in ("requests.log",):  # replaced by bridge.log; it was never trimmed
        try:
            os.remove(os.path.join(data_dir, stale))
        except OSError:
            pass
    return log


class LogWriter:
    """Stand-in for stdout/stderr when there is no console (windowed exe): sends print() output to the log."""

    is_log_writer = True

    def __init__(self):
        self.buf = ""

    def write(self, text):
        self.buf += text
        while "\n" in self.buf:
            line, self.buf = self.buf.split("\n", 1)
            if line.strip():
                log.info(line.rstrip())
        return len(text)

    def flush(self):
        pass

    def isatty(self):
        return False


def cleanup_cache(dl_dir, max_bytes, in_use):
    """Delete orphaned partial files, then the least recently used cached files until under max_bytes."""
    now = time.time()
    busy = in_use()
    cached = []
    total = 0
    try:
        names = os.listdir(dl_dir)
    except OSError:
        return
    for name in names:
        path = os.path.join(dl_dir, name)
        if name.endswith(".done"):
            if not os.path.exists(path[:-5]):
                _remove(path)
            continue
        try:
            st = os.stat(path)
        except OSError:
            continue
        if name.endswith(".upload"):
            if now - st.st_mtime > 86400:
                _remove(path)
            continue
        if path in busy:
            total += st.st_size
            continue
        done = path + ".done"
        if not os.path.exists(done):
            if now - st.st_mtime > 3600:  # abandoned partial download
                _remove(path)
            else:
                total += st.st_size
            continue
        cached.append((os.stat(done).st_mtime, st.st_size, path))
        total += st.st_size
    removed = 0
    for _used, size, path in sorted(cached):
        if total <= max_bytes:
            break
        _remove(path)
        _remove(path + ".done")
        total -= size
        removed += size
    if removed:
        log.info(f"Cache clean-up: freed {removed / 1048576:.0f} MB")


def _remove(path):
    try:
        os.remove(path)
    except OSError:
        pass


def start_cache_cleaner(sp, max_gb):
    max_bytes = int(max_gb * 1024 ** 3)

    def loop():
        while True:
            try:
                cleanup_cache(sp.dl_dir, max_bytes, sp.active_paths)
            except Exception as e:
                log.info(f"Cache clean-up failed: {e!r}")
            time.sleep(3600)

    threading.Thread(target=loop, daemon=True, name="cache-cleaner").start()
