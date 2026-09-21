import argparse
import json
from pathlib import Path

import pytest
from lxml import etree

from pptx_tool.__main__ import _load_theme, main
from pptx_tool.fix import (
    _has_monospace_font,
    _match_monospace_font,
    _normalize_slide_font,
    _update_first_level_bullet_style,
    _update_paragraph_style,
    fix_theme_font,
    local_tag,
    normalize_master_fonts,
    xmlns,
)
from pptx_tool.types import Theme

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
        ('<a:rPr><a:latin typeface="Pretendard"/><a:ea typeface="맑은 고딕"/></a:rPr>', False),
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
    # Stub out the pipeline so the real do_fix_pptx runs and we can see the Theme it builds.
    captured: list[Theme] = []
    monkeypatch.setattr('pptx_tool.__main__.extract_pptx', lambda src, dst: None)
    monkeypatch.setattr('pptx_tool.__main__.build_pptx', lambda src, dst: None)
    monkeypatch.setattr(
        'pptx_tool.__main__.fix_theme_font',
        lambda work_path, theme_info: captured.append(theme_info),
    )
    for func_name in ('normalize_master_fonts', 'normalize_layout_fonts', 'normalize_slide_fonts'):
        monkeypatch.setattr(f'pptx_tool.__main__.{func_name}', lambda work_path, theme_info: None)
    monkeypatch.setattr(
        'sys.argv',
        ['pptx-tool', 'fix-font', '--theme', str(theme_path), *argv, 'src.pptx', 'dst.pptx'],
    )
    main()
    assert [theme_info.preserve_mono for theme_info in captured] == [expected]


def _write_part(work_path: Path, relpath: str, body: str) -> Path:
    nsdecl = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in xmlns.items())
    part_path = work_path / relpath
    part_path.parent.mkdir(parents=True, exist_ok=True)
    part_path.write_text(f"<root {nsdecl}>{body}</root>")
    return part_path


@pytest.mark.parametrize(
    'preserve_mono,expected',
    [
        # The body style gets two passes: _update_paragraph_style turns Consolas into the
        # theme's monospace font, then _update_first_level_bullet_style re-reads the result.
        # The sentinel 'MONO-LT' is not a known monospace font, so the second pass falls back
        # to the minor font.  Every bundled theme has a monoFont that *is* a known monospace
        # font, so see test_master_body_style_double_pass for what they actually produce.
        (False, ["MIN-LT SemiBold", "MIN-SYM"]),
        (True, ["Consolas", "Consolas"]),
    ],
)
def test_normalize_master_fonts(tmp_path, make_theme, preserve_mono, expected):
    """The master's body style and its bullet fonts must honor the option too."""
    _write_part(tmp_path, 'ppt/presentation.xml', '<p:defaultTextStyle/>')
    master_path = _write_part(
        tmp_path,
        'ppt/slideMasters/slideMaster1.xml',
        '<p:bodyStyle>'
        '<a:lvl1pPr><a:defRPr><a:latin typeface="Consolas"/></a:defRPr></a:lvl1pPr>'
        '<a:lvl1pPr><a:buFont typeface="Consolas"/></a:lvl1pPr>'
        '</p:bodyStyle>',
    )
    theme_info = make_theme(body_first_level_style="SemiBold", preserve_mono=preserve_mono)
    normalize_master_fonts(tmp_path, theme_info)
    root_elem = etree.parse(master_path)
    typefaces = [
        elem.get('typeface')
        for elem in root_elem.xpath('//a:latin | //a:buFont', namespaces=xmlns)
    ]
    assert typefaces == expected


@pytest.mark.parametrize('preserve_mono', [False, True])
def test_normalize_master_fonts_still_applies_title_bold_to_preserved(tmp_path, make_theme, preserve_mono):
    """Preserving a typeface must not disable the title bold policy; weight is orthogonal."""
    master_path = _write_part(
        tmp_path,
        'ppt/slideMasters/slideMaster1.xml',
        '<p:titleStyle><a:defRPr><a:latin typeface="Consolas"/></a:defRPr></p:titleStyle>',
    )
    _write_part(tmp_path, 'ppt/presentation.xml', '<p:defaultTextStyle/>')
    normalize_master_fonts(tmp_path, make_theme(title_bold=True, preserve_mono=preserve_mono))
    root_elem = etree.parse(master_path)
    prop_elem = root_elem.xpath('//a:defRPr', namespaces=xmlns)[0]
    assert prop_elem.get('b') == '1'
    assert prop_elem.getchildren()[0].get('typeface') == ("Consolas" if preserve_mono else "MONO-LT")


@pytest.mark.parametrize('ph_type,expected', [("title", "+mj-lt"), ("body", "+mn-lt")])
def test_normalize_slide_font_placeholder_branch(make_theme, ph_type, expected):
    """A placeholder shape picks the major/minor scheme and also rewrites its defRPr."""
    root_elem = _parse_slide(
        f'<p:sp><p:nvSpPr><p:nvPr><p:ph type="{ph_type}"/></p:nvPr></p:nvSpPr><p:txBody><a:p>'
        '<a:pPr><a:defRPr><a:latin typeface="Pretendard"/></a:defRPr></a:pPr>'
        '<a:r><a:rPr><a:latin typeface="Consolas"/></a:rPr></a:r>'
        '</a:p></p:txBody></p:sp>'
    )
    _normalize_slide_font(root_elem, make_theme(preserve_mono=True), log_prefix="test")
    latin_elems = root_elem.xpath('//a:latin', namespaces=xmlns)
    assert [elem.get('typeface') for elem in latin_elems] == [expected, "Consolas"]


@pytest.mark.parametrize('preserve_mono', [False, True])
def test_fix_theme_font_is_never_preserved(tmp_path, make_theme, preserve_mono):
    """The theme font scheme defines the +mj-*/+mn-* placeholders, so it must always be replaced."""
    theme_path = _write_part(
        tmp_path,
        'ppt/theme/theme1.xml',
        '<a:fontScheme name="mono">'
        '<a:majorFont><a:latin typeface="Consolas"/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Consolas"/></a:minorFont>'
        '</a:fontScheme>',
    )
    fix_theme_font(tmp_path, make_theme(preserve_mono=preserve_mono))
    root_elem = etree.parse(theme_path)
    typefaces = [
        elem.get('typeface')
        for elem in root_elem.xpath('//a:majorFont/a:latin | //a:minorFont/a:latin', namespaces=xmlns)
    ]
    assert typefaces == ["MAJ-LT", "MIN-LT"]


def test_master_body_style_double_pass(tmp_path, make_theme):
    """A theme whose monoFont is itself a known monospace font keeps it in the body style.

    Every bundled theme is in this case: 7 of the 11 use 'Sarasa Term K', the rest use
    'NanumGothicCoding' or 'Consolas', and all of those are in known_monospace_fonts.  So the
    second pass over a first-level body style recognizes the font the first pass just wrote
    and keeps it, instead of falling back to the minor font.
    """
    master_path = _write_part(
        tmp_path,
        'ppt/slideMasters/slideMaster1.xml',
        '<p:bodyStyle><a:lvl1pPr><a:defRPr>'
        '<a:latin typeface="Consolas"/>'
        '</a:defRPr></a:lvl1pPr></p:bodyStyle>',
    )
    _write_part(tmp_path, 'ppt/presentation.xml', '<p:defaultTextStyle/>')
    theme_info = make_theme(mono_font_latin="Sarasa Term K", body_first_level_style="SemiBold")
    normalize_master_fonts(tmp_path, theme_info)
    root_elem = etree.parse(master_path)
    latin_elem = root_elem.xpath('//a:latin', namespaces=xmlns)[0]
    assert latin_elem.get('typeface') == "Sarasa Term K SemiBold"
