import dataclasses
from typing import Any

import pytest
from lxml import etree

from pptx_tool.fix import xmlns
from pptx_tool.types import Theme


@pytest.fixture
def make_theme():
    """Build a Theme with recognizable sentinel typefaces, overridable per test."""

    def _make_theme(**overrides: Any) -> Theme:
        theme_info = Theme(
            major_font_latin="MAJ-LT",
            major_font_hangul="MAJ-EA",
            major_font_symbol="MAJ-SYM",
            minor_font_latin="MIN-LT",
            minor_font_hangul="MIN-EA",
            minor_font_symbol="MIN-SYM",
            title_bold=True,
            mono_font_latin="MONO-LT",
            mono_font_hangul="MONO-EA",
        )
        return dataclasses.replace(theme_info, **overrides)

    return _make_theme


@pytest.fixture
def parse_fragment():
    """Parse an XML fragment with the 'a' and 'p' namespaces declared."""

    def _parse_fragment(xml: str) -> etree.Element:
        nsdecl = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in xmlns.items())
        return etree.fromstring(f"<root {nsdecl}>{xml}</root>")[0]

    return _parse_fragment
