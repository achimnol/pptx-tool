import contextlib
import os
import shutil
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

    When a reservation would exceed the quota, the working directories of the oldest finished
    requests are deleted first. The directories of the requests in progress are never deleted.
    Each server process needs its own root, as the eviction is coordinated only within a process.
    """

    def __init__(self, root: Path, quota: int) -> None:
        self.root = root
        self.quota = quota
        self._lock = threading.Lock()
        self._active: set[Path] = set()

    def _ensure_root(self) -> None:
        self.root.mkdir(parents=True, exist_ok=True, mode=0o700)
        if self.root.is_symlink():
            raise StorageError("The temporary directory must not be a symbolic link.")
        # A fixed path inside a shared /tmp may have been created by another local user in advance.
        if hasattr(os, "getuid") and self.root.stat().st_uid != os.getuid():
            raise StorageError("The temporary directory is owned by another user.")

    def cleanup_stale(self) -> None:
        """Delete everything left over from the previous runs."""
        self._ensure_root()
        with self._lock:
            for path in self.root.iterdir():
                if path not in self._active:
                    self._remove(path)

    @contextlib.contextmanager
    def request_dir(self) -> Iterator[Path]:
        """Create a working directory for a request, which is deleted when the request is done."""
        self._ensure_root()
        with self._lock:
            path = Path(tempfile.mkdtemp(prefix="req-", dir=self.root))
            self._active.add(path)
        try:
            yield path
        finally:
            with self._lock:
                self._active.discard(path)
            shutil.rmtree(path, ignore_errors=True)

    def reserve(self, needed: int) -> None:
        """Make room for the given number of bytes, evicting the oldest finished requests if needed.

        The contents already written by the requests in progress, including the caller's own,
        count toward the usage.
        """
        if needed > self.quota:
            raise RequestTooLargeError(f"The request needs more temporary storage than the limit ({self.quota} bytes).")
        with self._lock:
            entries = []
            for path in self.root.iterdir():
                try:
                    mtime = path.lstat().st_mtime
                except OSError:
                    continue
                entries.append((mtime, path.name, path, _tree_size(path) if path.is_dir() else path.lstat().st_size))
            used = sum(size for _, _, _, size in entries)
            for _, _, path, size in sorted(entries, key=lambda e: (e[0], e[1])):
                if used + needed <= self.quota:
                    return
                if path in self._active:
                    continue
                self._remove(path)
                used -= size
            if used + needed > self.quota:
                raise StorageFullError("The temporary storage is occupied by the requests in progress.")

    @staticmethod
    def _remove(path: Path) -> None:
        if path.is_dir() and not path.is_symlink():
            shutil.rmtree(path, ignore_errors=True)
        else:
            path.unlink(missing_ok=True)
