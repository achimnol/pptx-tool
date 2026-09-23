from pathlib import Path

import pytest
from lxml import etree

from pptx_tool.fix import (
    FontThemeError,
    FontThemeExistsError,
    InvalidFontThemeNameError,
    build_font_theme_xml,
    install_font_theme,
    validate_font_theme_name,
    xpath_elements,
)

from .conftest import MakeTheme


def test_build_font_theme_xml(make_theme: MakeTheme) -> None:
    root_elem = etree.fromstring(build_font_theme_xml(make_theme(), "나의 테마"))
    assert root_elem.get("name") == "나의 테마"
    assert [e.get("typeface") for e in xpath_elements(root_elem, "a:majorFont/*")] == [
        "MAJ-LT",
        "MAJ-EA",
        "MAJ-EA",
        "MAJ-SYM",
    ]
    assert [e.get("typeface") for e in xpath_elements(root_elem, "a:minorFont/*")] == [
        "MIN-LT",
        "MIN-EA",
        "MIN-EA",
        "MIN-SYM",
    ]


@pytest.mark.parametrize("name", ["My Theme", "나의 테마", "  padded  ", "a.b", "x" * 100])
def test_validate_font_theme_name_accepts(name: str) -> None:
    assert validate_font_theme_name(name) == name.strip()


@pytest.mark.parametrize(
    "name",
    [
        *("", "   ", ".", "..", "../evil", "ends."),
        *("a/b", "a\\b", "C:", "a*b", "a?b", 'a"b', "a<b", "a|b", "a\nb"),
        "x" * 101,
    ],
)
def test_validate_font_theme_name_rejects(name: str) -> None:
    with pytest.raises(InvalidFontThemeNameError):
        validate_font_theme_name(name)


def test_install_font_theme(tmp_path: Path) -> None:
    path = install_font_theme(b"<xml/>", "My Theme", theme_dir=tmp_path)
    assert path == tmp_path / "My Theme.xml"
    assert path.read_bytes() == b"<xml/>"
    with pytest.raises(FontThemeExistsError) as exc_info:
        install_font_theme(b"<new/>", "My Theme", theme_dir=tmp_path)
    assert exc_info.value.path == path
    assert path.read_bytes() == b"<xml/>"
    install_font_theme(b"<new/>", "My Theme", overwrite=True, theme_dir=tmp_path)
    assert path.read_bytes() == b"<new/>"


def test_install_font_theme_missing_dir(tmp_path: Path) -> None:
    with pytest.raises(FontThemeError, match="does not exist"):
        install_font_theme(b"<xml/>", "My Theme", theme_dir=tmp_path / "missing")


def test_install_font_theme_rejects_unsafe_name(tmp_path: Path) -> None:
    with pytest.raises(InvalidFontThemeNameError):
        install_font_theme(b"<xml/>", "../evil", theme_dir=tmp_path)
    assert not (tmp_path.parent / "evil.xml").exists()
