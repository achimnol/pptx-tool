import os
import stat
import threading
from pathlib import Path

import pytest

from pptx_tool.web.storage import RequestStorage, RequestTooLargeError, StorageError, StorageFullError


def _make_old_dir(root: Path, name: str, size: int, age: int) -> Path:
    path = root / name
    path.mkdir(parents=True)
    (path / "data").write_bytes(b"x" * size)
    mtime = path.stat().st_mtime - age
    os.utime(path, (mtime, mtime))
    return path


def test_request_dir_is_removed(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path / "a" / "b", quota=1000)
    with storage.request_dir() as req_dir:
        assert req_dir.parent == storage.root
        (req_dir / "file").write_bytes(b"x")
    assert storage.root.is_dir()
    assert not req_dir.exists()


def test_request_dir_is_removed_on_error(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path, quota=1000)
    with pytest.raises(RuntimeError), storage.request_dir() as req_dir:
        raise RuntimeError
    assert not req_dir.exists()


def test_cleanup_stale(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path, quota=1000)
    stale = _make_old_dir(tmp_path, "req-stale", 10, age=0)
    (tmp_path / "stray").write_bytes(b"x")
    storage.cleanup_stale()
    assert not stale.exists()
    assert not (tmp_path / "stray").exists()


def test_reserve_evicts_oldest_first(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path, quota=3000)
    old = _make_old_dir(tmp_path, "req-old", 1000, age=200)
    mid = _make_old_dir(tmp_path, "req-mid", 1000, age=100)
    with storage.request_dir() as req_dir:
        storage.reserve(req_dir, 1000)
        assert old.exists() and mid.exists()
        storage.reserve(req_dir, 2000)
        assert not old.exists() and mid.exists()
        storage.reserve(req_dir, 3000)
        assert not mid.exists()


def test_reserve_replaces_previous_reservation(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path, quota=3000)
    with storage.request_dir() as req_dir:
        storage.reserve(req_dir, 2000)
        # Not an increment: the second reservation is the whole footprint.
        storage.reserve(req_dir, 2500)
        (req_dir / "src.pptx").write_bytes(b"x" * 1000)
        # The written bytes are covered by the reservation, not added to it.
        storage.reserve(req_dir, 3000)


def test_reserve_rejects_oversized_request(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path, quota=1000)
    old = _make_old_dir(tmp_path, "req-old", 10, age=10)
    with storage.request_dir() as req_dir, pytest.raises(RequestTooLargeError):
        storage.reserve(req_dir, 1001)
    assert old.exists()


def test_reserve_counts_concurrent_reservations(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path, quota=3000)
    stale = _make_old_dir(tmp_path, "req-stale", 1500, age=10)
    with storage.request_dir() as first, storage.request_dir() as second:
        storage.reserve(first, 2000)
        assert not stale.exists()
        # Nothing on disk yet, but the first request's reservation is still counted.
        with pytest.raises(StorageFullError):
            storage.reserve(second, 2000)
        assert first.exists() and second.exists()
        storage.reserve(second, 1000)


def test_reserve_never_evicts_requests_in_progress(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path, quota=3000)
    started = threading.Event()
    release = threading.Event()

    def other_request() -> None:
        with storage.request_dir() as req_dir:
            (req_dir / "src.pptx").write_bytes(b"x" * 2000)
            started.set()
            release.wait()

    thread = threading.Thread(target=other_request)
    thread.start()
    try:
        assert started.wait(5)
        with storage.request_dir() as req_dir, pytest.raises(StorageFullError):
            storage.reserve(req_dir, 2000)
        assert list(storage.root.iterdir()) != []
    finally:
        release.set()
        thread.join()
    assert list(storage.root.iterdir()) == []


@pytest.mark.skipif(os.name != "posix", reason="creating symlinks may need privileges")
def test_root_must_not_be_symlink(tmp_path: Path) -> None:
    (tmp_path / "real").mkdir()
    (tmp_path / "link").symlink_to(tmp_path / "real")
    storage = RequestStorage(tmp_path / "link", quota=1000)
    with pytest.raises(StorageError):
        storage.cleanup_stale()


def test_root_must_be_a_directory(tmp_path: Path) -> None:
    (tmp_path / "file").write_bytes(b"x")
    storage = RequestStorage(tmp_path / "file", quota=1000)
    with pytest.raises(StorageError):
        storage.cleanup_stale()


@pytest.mark.skipif(os.name != "posix", reason="POSIX permissions")
def test_root_is_made_private(tmp_path: Path) -> None:
    root = tmp_path / "shared"
    root.mkdir(mode=0o755)
    RequestStorage(root, quota=1000).cleanup_stale()
    assert stat.S_IMODE(root.stat().st_mode) == 0o700
