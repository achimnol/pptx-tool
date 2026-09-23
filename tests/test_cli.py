import json
import logging
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pptx_tool.__main__ import main

THEME_DATA = {
    "majorFont": {"latin": "Major Sans", "hangul": "메이저", "symbol": "Major Symbol"},
    "minorFont": {"latin": "Minor Sans", "hangul": "마이너", "symbol": "Minor Symbol"},
    "monoFont": {"latin": "Mono Code", "hangul": "모노"},
    "options": {"titleBold": True, "bodyFirstLevelStyle": "SemiBold"},
}

# Kept as a list of lines since some of them end with a space.
FIX_FONT_OUTPUT_LINES = [
    "Current font scheme: (name='Office')",
    "  majorFont:",
    "    latin: Arial",
    "    ea: ",
    "    cs: ",
    "    font (Hang): 맑은 고딕",
    "  minorFont:",
    "    latin: Arial",
    "    ea: ",
    "    cs: ",
    "    font (Hang): 맑은 고딕",
    "New font scheme: (name='Office')",
    "  majorFont:",
    "    latin: Major Sans",
    "    ea: 메이저",
    "    cs: 메이저",
    "    sym: Major Symbol",
    "  minorFont:",
    "    latin: Minor Sans",
    "    ea: 마이너",
    "    cs: 마이너",
    "    sym: Minor Symbol",
    "Target monospace font:",
    "  latin: Mono Code",
    "  hangul: 모노",
    "slideLayout1.xml: template element (title)",
    "slide1.xml: normal element",
]


@pytest.fixture
def theme_path(tmp_path: Path) -> Path:
    path = tmp_path / "theme.json"
    path.write_text(json.dumps(THEME_DATA, ensure_ascii=False))
    return path


def _run_cli(monkeypatch: pytest.MonkeyPatch, *argv: str) -> None:
    monkeypatch.setattr("sys.argv", ["pptx-tool", *argv])
    main()


def test_fix_font_output(
    tmp_path: Path,
    sample_pptx: Path,
    theme_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    dst_path = tmp_path / "fixed.pptx"
    _run_cli(monkeypatch, "fix-font", "--theme", str(theme_path), str(sample_pptx), str(dst_path))
    assert capsys.readouterr().out == "\n".join(FIX_FONT_OUTPUT_LINES) + "\n"
    with zipfile.ZipFile(dst_path) as zf:
        theme_xml = zf.read("ppt/theme/theme1.xml").decode()
        slide_xml = zf.read("ppt/slides/slide1.xml").decode()
    assert 'typeface="Major Sans"' in theme_xml
    assert 'typeface="+mn-lt"' in slide_xml
    assert 'typeface="Mono Code"' in slide_xml


def test_fix_font_output_preserve_mono(
    tmp_path: Path,
    sample_pptx: Path,
    theme_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    dst_path = tmp_path / "fixed.pptx"
    _run_cli(monkeypatch, "fix-font", "--theme", str(theme_path), "--preserve-mono", str(sample_pptx), str(dst_path))
    expected_lines = [
        *FIX_FONT_OUTPUT_LINES[:22],
        "Preserving the existing monospace fonts as-is.",
        *FIX_FONT_OUTPUT_LINES[25:],
    ]
    assert capsys.readouterr().out == "\n".join(expected_lines) + "\n"
    with zipfile.ZipFile(dst_path) as zf:
        slide_xml = zf.read("ppt/slides/slide1.xml").decode()
    assert 'typeface="Consolas"' in slide_xml


def test_generate_font_theme_output(
    tmp_path: Path,
    theme_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    theme_dir = tmp_path / "Theme Fonts"
    theme_dir.mkdir()
    monkeypatch.setattr("pptx_tool.fix._get_font_theme_dir", lambda: theme_dir)
    _run_cli(monkeypatch, "generate-font-theme", "--theme", str(theme_path), "My Theme")
    xml_path = theme_dir / "My Theme.xml"
    assert capsys.readouterr().out == f"Stored an Office theme font definition at:\n{xml_path}\n"
    assert xml_path.read_bytes() == (
        b'<a:fontScheme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="My Theme">\n'
        b"  <a:majorFont>\n"
        b'    <a:latin typeface="Major Sans"/>\n'
        b'    <a:ea typeface="&#47700;&#51060;&#51200;"/>\n'
        b'    <a:cs typeface="&#47700;&#51060;&#51200;"/>\n'
        b'    <a:sym typeface="Major Symbol"/>\n'
        b"  </a:majorFont>\n"
        b"  <a:minorFont>\n"
        b'    <a:latin typeface="Minor Sans"/>\n'
        b'    <a:ea typeface="&#47560;&#51060;&#45320;"/>\n'
        b'    <a:cs typeface="&#47560;&#51060;&#45320;"/>\n'
        b'    <a:sym typeface="Minor Symbol"/>\n'
        b"  </a:minorFont>\n"
        b"</a:fontScheme>\n"
    )


def test_invalid_theme_exits_with_error(
    tmp_path: Path,
    sample_pptx: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    theme_path = tmp_path / "theme.json"
    theme_path.write_text(json.dumps({**THEME_DATA, "majorFont": {"latin": "A"}}))
    with pytest.raises(SystemExit) as exc_info:
        _run_cli(monkeypatch, "fix-font", "--theme", str(theme_path), str(sample_pptx), str(tmp_path / "out.pptx"))
    assert exc_info.value.code == 2
    assert "majorFont.hangul: must be a non-empty string" in capsys.readouterr().err


def test_fix_font_with_bundled_theme(tmp_path: Path, sample_pptx: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    dst_path = tmp_path / "fixed.pptx"
    _run_cli(monkeypatch, "fix-font", "--theme", "pretendard", str(sample_pptx), str(dst_path))
    with zipfile.ZipFile(dst_path) as zf:
        assert 'typeface="Pretendard"' in zf.read("ppt/theme/theme1.xml").decode()


LOOPBACK_HOSTS = ("127.0.0.1", "localhost", "[::1]")


def test_serve_without_web_extra(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setitem(sys.modules, "uvicorn", None)
    with pytest.raises(SystemExit) as exc_info:
        _run_cli(monkeypatch, "serve")
    assert "requires the 'web' extra" in str(exc_info.value.code)


def test_serve_refuses_local_on_public_address(monkeypatch: pytest.MonkeyPatch) -> None:
    pytest.importorskip("litestar")
    with pytest.raises(SystemExit) as exc_info:
        _run_cli(monkeypatch, "serve", "--host", "0.0.0.0", "--local")
    assert "loopback" in str(exc_info.value.code)


@pytest.mark.parametrize(
    "argv,local,allowed_hosts",
    [
        ([], False, (*LOOPBACK_HOSTS, "127.0.0.1:8000", "localhost:8000", "[::1]:8000")),
        (["--local", "--port", "9000"], True, (*LOOPBACK_HOSTS, "127.0.0.1:9000", "localhost:9000", "[::1]:9000")),
        (["--host", "0.0.0.0"], False, ()),
    ],
)
def test_serve_config(
    monkeypatch: pytest.MonkeyPatch, argv: list[str], local: bool, allowed_hosts: tuple[str, ...]
) -> None:
    pytest.importorskip("litestar")
    import uvicorn

    from pptx_tool.web import app as web_app

    captured: dict[str, Any] = {}
    package_logger = logging.getLogger("pptx_tool")
    monkeypatch.setattr(package_logger, "propagate", True)
    monkeypatch.setattr(web_app, "create_app", lambda web_config: captured.setdefault("config", web_config))
    monkeypatch.setattr(uvicorn, "run", lambda app, **kwargs: captured.update(kwargs))
    _run_cli(monkeypatch, "serve", "--max-upload-mb", "10", *argv)
    assert captured["config"].local is local
    assert captured["config"].allowed_hosts == allowed_hosts
    assert captured["config"].max_upload_size == 10 * 1024 * 1024
    # The server does not echo the processing logs of each request.
    assert package_logger.propagate is False


def test_generate_font_theme_exists(
    tmp_path: Path, theme_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    monkeypatch.setattr("pptx_tool.fix._get_font_theme_dir", lambda: tmp_path)
    (tmp_path / "My Theme.xml").write_text("")
    with pytest.raises(SystemExit) as exc_info:
        _run_cli(monkeypatch, "generate-font-theme", "--theme", str(theme_path), "My Theme")
    assert "The target theme file already exist." in str(exc_info.value.code)
