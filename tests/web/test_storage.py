import os
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
    storage.reserve(1000)
    assert old.exists() and mid.exists()
    storage.reserve(2000)
    assert not old.exists() and mid.exists()
    storage.reserve(3000)
    assert not mid.exists()


def test_reserve_rejects_oversized_request(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path, quota=1000)
    old = _make_old_dir(tmp_path, "req-old", 10, age=10)
    with pytest.raises(RequestTooLargeError):
        storage.reserve(1001)
    assert old.exists()


def test_reserve_keeps_active_dirs(tmp_path: Path) -> None:
    storage = RequestStorage(tmp_path, quota=3000)
    stale = _make_old_dir(tmp_path, "req-stale", 1000, age=10)
    with storage.request_dir() as req_dir:
        (req_dir / "src.pptx").write_bytes(b"x" * 1500)
        with pytest.raises(StorageFullError):
            storage.reserve(2000)
        assert not stale.exists()
        assert req_dir.exists()


@pytest.mark.skipif(not hasattr(os, "getuid"), reason="POSIX only")
def test_root_must_not_be_symlink(tmp_path: Path) -> None:
    (tmp_path / "real").mkdir()
    (tmp_path / "link").symlink_to(tmp_path / "real")
    storage = RequestStorage(tmp_path / "link", quota=1000)
    with pytest.raises(StorageError):
        storage.cleanup_stale()
