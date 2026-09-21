import argparse
import json
from pathlib import Path

import pytest
from lxml import etree

from pptx_tool.__main__ import _load_theme, _resolve_preserve_mono, main
from pptx_tool.fix import (
    _has_monospace_font,
    _match_monospace_font,
    _normalize_slide_font,
    _update_first_level_bullet_style,
    _update_paragraph_style,
    local_tag,
    xmlns,
)

THEMES_DIR = Path(__file__).parent.parent / 'themes'


def _parse_slide(body: str) -> etree.ElementTree:
    nsdecl = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in xmlns.items())
    return etree.ElementTree(etree.fromstring(f"<p:sld {nsdecl}>{body}</p:sld>"))


@pytest.mark.parametrize(
    'typeface,expected',
    [
        ("Consolas", True),
        ("CONSOLAS", True),
        ("Fira Code", True),
        ("Sarasa Term K", True),
        ("Victor Mono", True),
        ("PragmataPro", True),
        ("Pretendard", False),
        ("", False),
        (None, False),
    ],
)
def test_match_monospace_font(typeface, expected):
    assert _match_monospace_font(typeface) is expected


@pytest.mark.parametrize(
    'fragment,expected',
    [
        ('<a:rPr><a:latin typeface="Consolas"/></a:rPr>', True),
        ('<a:rPr><a:latin typeface="Pretendard"/><a:ea typeface="D2Coding"/></a:rPr>', False),
        ('<a:rPr><a:latin typeface="Pretendard"/><a:ea typeface="NanumGothicCoding"/></a:rPr>', True),
        ('<a:rPr><a:cs typeface="Menlo"/></a:rPr>', True),
        ('<a:rPr><a:sym typeface="SF Mono"/></a:rPr>', True),
        ('<a:rPr><a:font script="Hang" typeface="Hack"/></a:rPr>', True),
        ('<a:rPr><a:latin typeface="Pretendard"/><a:ea typeface="Pretendard"/></a:rPr>', False),
        ('<a:rPr/>', False),
        # A child without a typeface attribute must not raise.
        ('<a:rPr><a:sym/><a:latin typeface="Pretendard"/></a:rPr>', False),
        # A non-typeface child carrying a typeface-like name must be ignored.
        ('<a:rPr><a:highlight typeface="Consolas"/></a:rPr>', False),
    ],
)
def test_has_monospace_font(parse_fragment, fragment, expected):
    assert _has_monospace_font(parse_fragment(fragment)) is expected


MIXED_RUN = (
    '<a:rPr sz="1800" dirty="0">'
    '<a:latin typeface="Consolas"/>'
    '<a:ea typeface="굴림"/>'
    '<a:cs typeface="Consolas"/>'
    '<a:sym typeface="Wingdings"/>'
    '<a:font script="Hang" typeface="굴림"/>'
    '</a:rPr>'
)


def test_paragraph_style_converts_monospace_when_not_preserving(parse_fragment, make_theme):
    """The default path must keep replacing monospace fonts with the theme's monospace font."""
    prop_elem = parse_fragment(MIXED_RUN)
    _update_paragraph_style(prop_elem, make_theme())
    assert dict(prop_elem.attrib) == {'sz': "1800", 'dirty': "0"}
    # The 'font' child is dropped and every other typeface is rewritten.
    assert [(local_tag(elem), dict(elem.attrib)) for elem in prop_elem.getchildren()] == [
        ("latin", {'typeface': "MONO-LT"}),
        ("ea", {'typeface': "+mn-ea"}),
        ("cs", {'typeface': "MONO-EA"}),
        ("sym", {'typeface': "MIN-SYM"}),
    ]


def test_paragraph_style_preserves_monospace_run_verbatim(parse_fragment, make_theme):
    """A monospace run must stay byte-identical, including its sym child and extra attributes."""
    prop_elem = parse_fragment(MIXED_RUN)
    before = etree.tostring(prop_elem)
    _update_paragraph_style(prop_elem, make_theme(preserve_mono=True))
    assert etree.tostring(prop_elem) == before


@pytest.mark.parametrize('scheme_prefix,expected_sym', [("mn", "MIN-SYM"), ("mj", "MAJ-SYM")])
def test_paragraph_style_still_converts_non_monospace_when_preserving(
    parse_fragment, make_theme, scheme_prefix, expected_sym,
):
    prop_elem = parse_fragment(
        '<a:rPr>'
        '<a:latin typeface="Pretendard"/>'
        '<a:ea typeface="Pretendard"/>'
        '<a:cs typeface="Pretendard"/>'
        '<a:sym typeface="Wingdings"/>'
        '<a:font script="Hang" typeface="Pretendard"/>'
        '</a:rPr>'
    )
    _update_paragraph_style(prop_elem, make_theme(preserve_mono=True), scheme_prefix=scheme_prefix)
    typefaces = [elem.get('typeface') for elem in prop_elem.getchildren()]
    assert typefaces == [
        f"+{scheme_prefix}-lt",
        f"+{scheme_prefix}-ea",
        f"+{scheme_prefix}-cs",
        expected_sym,
    ]


@pytest.mark.parametrize(
    'preserve_mono,fragment,expected',
    [
        (False, '<a:defRPr><a:latin typeface="Consolas"/></a:defRPr>', "MONO-LT SemiBold"),
        (False, '<a:defRPr><a:latin typeface="Pretendard"/></a:defRPr>', "MIN-LT SemiBold"),
        (True, '<a:defRPr><a:latin typeface="Consolas"/></a:defRPr>', "Consolas"),
        (True, '<a:defRPr><a:latin typeface="Pretendard"/></a:defRPr>', "MIN-LT SemiBold"),
        (False, '<a:defRPr><a:ea typeface="NanumGothicCoding"/></a:defRPr>', "MONO-EA SemiBold"),
        (True, '<a:defRPr><a:ea typeface="NanumGothicCoding"/></a:defRPr>', "NanumGothicCoding"),
    ],
)
def test_first_level_bullet_style(parse_fragment, make_theme, preserve_mono, fragment, expected):
    prop_elem = parse_fragment(fragment)
    theme_info = make_theme(body_first_level_style="SemiBold", preserve_mono=preserve_mono)
    _update_first_level_bullet_style(prop_elem, theme_info)
    assert prop_elem.getchildren()[0].get('typeface') == expected


@pytest.mark.parametrize(
    'preserve_mono,expected',
    [
        (False, ["MIN-SYM", "MIN-SYM", "MIN-SYM"]),
        (True, ["Consolas", "MIN-SYM", "MIN-SYM"]),
    ],
)
def test_normalize_slide_font_bullet_fonts(make_theme, preserve_mono, expected):
    root_elem = _parse_slide(
        '<a:pPr><a:buFont typeface="Consolas"/></a:pPr>'
        '<a:pPr><a:buFont typeface="Wingdings"/></a:pPr>'
        '<a:pPr><a:buFont charset="2"/></a:pPr>'  # no typeface: must not raise
    )
    _normalize_slide_font(root_elem, make_theme(preserve_mono=preserve_mono), log_prefix="test")
    bullet_fonts = root_elem.xpath('//a:buFont', namespaces=xmlns)
    assert [elem.get('typeface') for elem in bullet_fonts] == expected


def test_normalize_slide_font_preserves_monospace_shape(make_theme):
    root_elem = _parse_slide(
        '<p:sp><p:txBody><a:p>'
        '<a:r><a:rPr><a:latin typeface="JetBrains Mono"/></a:rPr></a:r>'
        '<a:r><a:rPr><a:latin typeface="Pretendard"/></a:rPr></a:r>'
        '</a:p></p:txBody></p:sp>'
    )
    _normalize_slide_font(root_elem, make_theme(preserve_mono=True), log_prefix="test")
    latin_elems = root_elem.xpath('//a:latin', namespaces=xmlns)
    assert [elem.get('typeface') for elem in latin_elems] == ["JetBrains Mono", "+mn-lt"]


@pytest.mark.parametrize('theme_path', sorted(THEMES_DIR.glob('*.json')), ids=lambda p: p.name)
def test_bundled_themes_load_without_preserve_mono_key(theme_path):
    """Themes that omit 'preserveMono' must keep loading, with the option off."""
    theme_info = _load_theme(argparse.Namespace(theme=theme_path))
    assert theme_info.preserve_mono is False


@pytest.mark.parametrize(
    'argv,expected',
    [
        ([], True),
        (['--preserve-mono'], True),
        (['--no-preserve-mono'], False),
    ],
)
def test_cli_overrides_theme_preserve_mono(tmp_path, monkeypatch, argv, expected):
    theme_path = tmp_path / 'theme.json'
    theme_path.write_text(json.dumps({
        'majorFont': {'latin': 'A', 'hangul': 'A', 'symbol': 'A'},
        'minorFont': {'latin': 'A', 'hangul': 'A', 'symbol': 'A'},
        'monoFont': {'latin': 'M', 'hangul': 'M'},
        'options': {'titleBold': True, 'bodyFirstLevelStyle': None, 'preserveMono': True},
    }))
    captured: list[bool] = []
    monkeypatch.setattr(
        'pptx_tool.__main__.do_fix_pptx',
        lambda args: captured.append(_resolve_preserve_mono(_load_theme(args), args).preserve_mono),
    )
    monkeypatch.setattr(
        'sys.argv',
        ['pptx-tool', 'fix-font', '--theme', str(theme_path), *argv, 'src.pptx', 'dst.pptx'],
    )
    main()
    assert captured == [expected]
