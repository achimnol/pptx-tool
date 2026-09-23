import contextlib
import os
import shutil
import stat
import tempfile
import threading
from collections.abc import Iterator
from pathlib import Path


class StorageError(Exception):
    """Raised when the temporary storage cannot serve a request. Messages must not contain server paths."""


class RequestTooLargeError(StorageError):
    """A single request needs more space than the whole quota."""


class StorageFullError(StorageError):
    """The requests in progress occupy the quota, so nothing can be evicted."""


def _tree_size(path: Path) -> int:
    """The total size of the files under path, tolerating entries that vanish while scanning."""
    total = 0
    try:
        with os.scandir(path) as entries:
            for entry in entries:
                try:
                    if entry.is_dir(follow_symlinks=False):
                        total += _tree_size(Path(entry.path))
                    else:
                        total += entry.stat(follow_symlinks=False).st_size
                except OSError:
                    continue
    except OSError:
        pass
    return total


class RequestStorage:
    """A fixed directory holding one working directory per request, bounded by a total size quota.

    Each request reserves the space it is going to use, so that concurrent requests cannot exceed
    the quota together. When a reservation would exceed the quota, the leftovers of the finished
    requests are deleted first, oldest first. The directories of the requests in progress are never
    deleted. Each server process needs its own root, as the accounting is done only within a process.
    """

    def __init__(self, root: Path, quota: int) -> None:
        self.root = root
        self.quota = quota
        self._lock = threading.Lock()
        self._reserved: dict[Path, int] = {}
        """The reserved footprint of each request in progress, keyed by its working directory."""

    def _ensure_root(self) -> None:
        try:
            self.root.mkdir(parents=True, exist_ok=True, mode=0o700)
        except OSError as e:
            raise StorageError("The temporary directory cannot be created.") from e
        if self.root.is_symlink():
            raise StorageError("The temporary directory must not be a symbolic link.")
        # A fixed path inside a shared /tmp may have been created by another local user in advance.
        st = self.root.stat()
        if hasattr(os, "getuid") and st.st_uid != os.getuid():
            raise StorageError("The temporary directory is owned by another user.")
        # The mode given to mkdir() is subject to the umask and ignored for an existing directory.
        if os.name == "posix" and stat.S_IMODE(st.st_mode) != 0o700:
            self.root.chmod(0o700)

    def cleanup_stale(self) -> None:
        """Delete everything left over from the previous runs."""
        self._ensure_root()
        with self._lock:
            for path in self.root.iterdir():
                if path not in self._reserved:
                    self._remove(path)

    @contextlib.contextmanager
    def request_dir(self) -> Iterator[Path]:
        """Create a working directory for a request, which is deleted when the request is done."""
        self._ensure_root()
        with self._lock:
            path = Path(tempfile.mkdtemp(prefix="req-", dir=self.root))
            self._reserved[path] = 0
        try:
            yield path
        finally:
            shutil.rmtree(path, ignore_errors=True)
            with self._lock:
                del self._reserved[path]

    def reserve(self, request_dir: Path, needed: int) -> None:
        """Reserve the total footprint of a request, evicting the leftovers of finished requests if needed.

        The reservation replaces the previous one of the same request, so callers pass the whole
        footprint they expect so far, not an increment.
        """
        if needed > self.quota:
            raise RequestTooLargeError(f"The request needs more temporary storage than the limit ({self.quota} bytes).")
        with self._lock:
            self._reserved[request_dir] = needed
            entries = []
            for path in self.root.iterdir():
                try:
                    st = path.lstat()
                except OSError:
                    continue
                size = _tree_size(path) if stat.S_ISDIR(st.st_mode) else st.st_size
                # A request in progress may not have written its reservation yet.
                entries.append((st.st_mtime, path.name, path, max(size, self._reserved.get(path, 0))))
            used = sum(size for _, _, _, size in entries)
            # A finished request's directory keeps the mtime of its last top-level entry, i.e. near its end.
            for _, _, path, size in sorted(entries, key=lambda e: (e[0], e[1])):
                if used <= self.quota:
                    return
                if path in self._reserved:
                    continue
                self._remove(path)
                used -= size
            if used > self.quota:
                raise StorageFullError("The temporary storage is occupied by the requests in progress.")

    @staticmethod
    def _remove(path: Path) -> None:
        if path.is_dir() and not path.is_symlink():
            shutil.rmtree(path, ignore_errors=True)
        else:
            path.unlink(missing_ok=True)
