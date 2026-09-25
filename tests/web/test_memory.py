"""Check that the server streams large files instead of holding them in the memory.

The server runs in a subprocess with the same options as in production, and its peak resident set
size is measured around a round trip of a pptx file several times larger than the memory budget.
Any step that holds the whole file in the memory, such as reading the upload or the result at once,
exceeds the budget.
"""

import json
import os
import socket
import subprocess
import sys
import time
import zipfile
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest

from ..samples import make_minimal_pptx

PPTX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MEDIA_SIZE = 128 * 1024**2
"""The size of the incompressible media file in the large pptx."""
MEMORY_BUDGET = 32 * 1024**2
"""The allowed growth of the server's peak resident set size while handling the large pptx."""
THEME = {
    "majorFont": {"latin": "Major Sans", "hangul": "메이저", "symbol": "Major Symbol"},
    "minorFont": {"latin": "Minor Sans", "hangul": "마이너", "symbol": "Minor Symbol"},
    "monoFont": {"latin": "Mono Code", "hangul": "모노"},
    "options": {"titleBold": True, "bodyFirstLevelStyle": "SemiBold", "preserveMono": False},
}

pytestmark = pytest.mark.skipif(
    not os.access("/proc/self/clear_refs", os.W_OK),
    reason="Resetting the peak resident set size requires Linux /proc/<pid>/clear_refs.",
)


def _free_port() -> int:
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _proc_status_kib(pid: int, key: str) -> int:
    for line in Path(f"/proc/{pid}/status").read_text().splitlines():
        if line.startswith(f"{key}:"):
            return int(line.split()[1])
    raise KeyError(key)


def _reset_peak_rss(pid: int) -> int:
    """Reset the peak resident set size (VmHWM) to the current one, and return it in bytes."""
    Path(f"/proc/{pid}/clear_refs").write_text("5")
    return _proc_status_kib(pid, "VmRSS") * 1024


def _peak_rss(pid: int) -> int:
    return _proc_status_kib(pid, "VmHWM") * 1024


def _make_large_pptx(path: Path, media_size: int) -> Path:
    make_minimal_pptx(path)
    # Random bytes cannot be compressed, so the file travels at its full size in every step.
    with zipfile.ZipFile(path, "a") as zf, zf.open("ppt/media/media1.bin", "w", force_zip64=True) as f:
        for _ in range(media_size // (1024 * 1024)):
            f.write(os.urandom(1024 * 1024))
    return path


@pytest.fixture
def server(tmp_path: Path) -> Iterator[tuple[subprocess.Popen[bytes], str]]:
    port = _free_port()
    proc = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "pptx_tool",
            "serve",
            "--port",
            str(port),
            "--tmp-dir",
            str(tmp_path / "server-tmp"),
            "--max-upload-mb",
            str(MEDIA_SIZE // 1024**2 + 16),
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
    )
    base_url = f"http://127.0.0.1:{port}"
    try:
        deadline = time.monotonic() + 30
        while True:
            try:
                httpx.get(f"{base_url}/api/config").raise_for_status()
                break
            except httpx.TransportError:
                if proc.poll() is not None or time.monotonic() > deadline:
                    stderr = proc.stderr.read().decode() if proc.stderr else ""
                    pytest.fail(f"The server did not start:\n{stderr}")
                time.sleep(0.1)
        yield proc, base_url
    finally:
        proc.terminate()
        proc.wait(timeout=10)


def _fix_font(client: httpx.Client, src: Path, dst: Path) -> None:
    with src.open("rb") as f:
        resp = client.post(
            "/api/fix-font",
            files={"file": (src.name, f, PPTX_MEDIA_TYPE)},
            data={"theme": json.dumps(THEME)},
        )
    assert resp.status_code == 200, resp.text
    with client.stream("GET", resp.json()["downloadUrl"]) as download:
        assert download.status_code == 200
        with dst.open("wb") as f:
            for chunk in download.iter_bytes():
                f.write(chunk)


def test_fix_font_streams_large_files(server: tuple[subprocess.Popen[bytes], str], tmp_path: Path) -> None:
    proc, base_url = server
    large = _make_large_pptx(tmp_path / "large.pptx", MEDIA_SIZE)
    assert large.stat().st_size > 4 * MEMORY_BUDGET
    with httpx.Client(base_url=base_url, timeout=300) as client:
        # Load the lazily imported modules and warm up the allocator with a small file first.
        _fix_font(client, make_minimal_pptx(tmp_path / "small.pptx"), tmp_path / "small-fixed.pptx")
        baseline = _reset_peak_rss(proc.pid)
        _fix_font(client, large, tmp_path / "large-fixed.pptx")
    growth = _peak_rss(proc.pid) - baseline
    assert growth < MEMORY_BUDGET, (
        f"The server's peak RSS grew by {growth / 1024**2:.1f} MiB for a {large.stat().st_size / 1024**2:.0f} MiB"
        f" pptx, over the budget of {MEMORY_BUDGET / 1024**2:.0f} MiB."
    )
    with zipfile.ZipFile(tmp_path / "large-fixed.pptx") as zf:
        assert zf.getinfo("ppt/media/media1.bin").file_size == MEDIA_SIZE
