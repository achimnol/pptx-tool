import zipfile
from pathlib import Path

import pytest

from pptx_tool.package import ArchiveLimits, UnsafeArchiveError, build_pptx, extract_pptx


def test_extract_and_build_round_trip(tmp_path: Path, sample_pptx: Path) -> None:
    work_dir = tmp_path / "work"
    extract_pptx(sample_pptx, work_dir)
    rebuilt = tmp_path / "rebuilt.pptx"
    build_pptx(work_dir, rebuilt)
    with zipfile.ZipFile(sample_pptx) as a, zipfile.ZipFile(rebuilt) as b:
        assert sorted(a.namelist()) == sorted(b.namelist())
        for name in a.namelist():
            assert a.read(name) == b.read(name)


def test_extract_accepts_file_object(tmp_path: Path, sample_pptx: Path) -> None:
    with sample_pptx.open("rb") as f:
        extract_pptx(f, tmp_path / "work", limits=ArchiveLimits())
    assert (tmp_path / "work" / "ppt" / "presentation.xml").is_file()


def test_extract_rejects_too_many_entries(tmp_path: Path, sample_pptx: Path) -> None:
    with pytest.raises(UnsafeArchiveError, match="too many entries"):
        extract_pptx(sample_pptx, tmp_path / "work", limits=ArchiveLimits(max_entries=3))
    assert not (tmp_path / "work").exists()


def test_extract_rejects_too_large_content(tmp_path: Path) -> None:
    bomb = tmp_path / "bomb.pptx"
    with zipfile.ZipFile(bomb, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("ppt/presentation.xml", b"\0" * 1_000_000)
    assert bomb.stat().st_size < 10_000
    with pytest.raises(UnsafeArchiveError, match="too large"):
        extract_pptx(bomb, tmp_path / "work", limits=ArchiveLimits(max_total_size=100_000))
