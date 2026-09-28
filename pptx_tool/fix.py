import logging
import re
import sys
from pathlib import Path
from typing import Final, NamedTuple, cast

from lxml import etree

from .types import Theme

logger = logging.getLogger(__name__)

xmlns: Final = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "dsp": "http://schemas.microsoft.com/office/drawing/2008/diagram",
}

style_typeface_map: Final = {
    "latin": "lt",
    "ea": "ea",
    "cs": "cs",
    "sym": "sym",
}

known_monospace_fonts: Final = {
    "courier",
    "courier new",
    "consolas",
    "cascadia code",
    "cascadia mono",
    "inconsolata",
    "jetbrains mono",
    "menlo",
    "monaco",
    "terminus",
    "bitstream vera sans mono",
    "droid sans mono",
    "dejavu sans mono",
    "pt mono",
    "sf mono",
    "andale mono",
    "arpercu mono",
    "dank mono",
    "input mono",
    "roboto mono",
    "oxygen mono",
    "space mono",
    "ubuntu mono",
    "liberation mono",
    "anonymous pro",
    "source code pro",
    "iosevka",
    "monolisa",
    "monoid",
    "gintronic",
    "fira code",
    "nanumgothiccoding",
    "hack",
    "recursive",
    "sarasa term k",
    "victor mono",
    "pragmatapro",
}

preservable_typeface_tags: Final = frozenset({"latin", "ea", "cs", "sym", "font"})

# The weight and slope suffixes of a typeface name such as "Pretendard ExtraBold" or "Segoe UI Semi Bold Italic".
# "Book" is left out because it is a part of family names like "Franklin Gothic Book".
_variant_suffix_pattern: Final = re.compile(
    r"(?:\s+(?:(?P<prefix>extra|ultra|semi|demi)[\s-]?)?"
    r"(?P<weight>thin|hairline|light|regular|normal|medium|bold|black|heavy))?"
    r"(?:\s+(?P<slope>italic|oblique))?$",
    re.IGNORECASE,
)


def local_tag(tag_name: str | bytes | etree._Element) -> str:
    """Strip out the namespace from the element tag name."""
    return etree.QName(tag_name).localname


def xpath_elements(node: etree._Element | etree._ElementTree, path: str) -> list[etree._Element]:
    """Evaluate an element-selecting XPath with the package namespaces bound."""
    return cast(list[etree._Element], node.xpath(path, namespaces=xmlns))


def _print_font_scheme(font_scheme: etree._Element, indent: str = "") -> None:
    for font_elem in font_scheme:
        logger.info("%s%s:", indent, local_tag(font_elem.tag))
        for prev_typeface in font_elem:
            prev_script_name = local_tag(prev_typeface.tag)
            match prev_script_name:
                case "font":
                    logger.info(
                        "%s  %s (%s): %s",
                        indent,
                        prev_script_name,
                        prev_typeface.get("script"),
                        prev_typeface.get("typeface"),
                    )
                case _:
                    logger.info("%s  %s: %s", indent, prev_script_name, prev_typeface.get("typeface"))


def _fill_font_scheme(target_elem: etree._Element, theme_info: Theme) -> None:
    major_font_elem = etree.Element(etree.QName(xmlns["a"], "majorFont"))
    major_font_elem.append(
        etree.Element(etree.QName(xmlns["a"], "latin"), attrib={"typeface": theme_info.major_font_latin})
    )
    major_font_elem.append(
        etree.Element(etree.QName(xmlns["a"], "ea"), attrib={"typeface": theme_info.major_font_hangul})
    )
    major_font_elem.append(
        etree.Element(etree.QName(xmlns["a"], "cs"), attrib={"typeface": theme_info.major_font_hangul})
    )
    major_font_elem.append(
        etree.Element(etree.QName(xmlns["a"], "sym"), attrib={"typeface": theme_info.major_font_symbol})
    )
    # major_font_elem.append(
    #     etree.Element(
    #         etree.QName(xmlns['a'], 'font'),
    #         attrib={'script': 'Hang', 'typeface': theme_info.major_font_hangul},
    #     )
    # )
    minor_font_elem = etree.Element(etree.QName(xmlns["a"], "minorFont"))
    minor_font_elem.append(
        etree.Element(etree.QName(xmlns["a"], "latin"), attrib={"typeface": theme_info.minor_font_latin})
    )
    minor_font_elem.append(
        etree.Element(etree.QName(xmlns["a"], "ea"), attrib={"typeface": theme_info.minor_font_hangul})
    )
    minor_font_elem.append(
        etree.Element(etree.QName(xmlns["a"], "cs"), attrib={"typeface": theme_info.minor_font_hangul})
    )
    minor_font_elem.append(
        etree.Element(etree.QName(xmlns["a"], "sym"), attrib={"typeface": theme_info.minor_font_symbol})
    )
    # minor_font_elem.append(
    #     etree.Element(
    #         etree.QName(xmlns['a'], 'font'),
    #         attrib={'script': 'Hang', 'typeface': theme_info.minor_font_hangul},
    #     )
    # )
    # Replace the theme font
    font_scheme_attrib = dict(target_elem.items())
    target_elem.clear()
    for k, v in font_scheme_attrib.items():
        target_elem.set(k, v)
    target_elem.append(major_font_elem)
    target_elem.append(minor_font_elem)


def fix_theme_font(
    work_path: Path,
    theme_info: Theme,
) -> None:
    theme_dir = work_path / "ppt" / "theme"
    for theme_path in theme_dir.glob("theme*.xml"):
        root_elem = etree.parse(theme_path)
        font_scheme_elem = xpath_elements(root_elem, "//a:fontScheme")[0]

        # Print out current theme font configuration
        logger.info("Current font scheme: (name=%r)", font_scheme_elem.get("name"))
        _print_font_scheme(font_scheme_elem, indent="  ")

        _fill_font_scheme(font_scheme_elem, theme_info)

        logger.info("New font scheme: (name=%r)", font_scheme_elem.get("name"))
        _print_font_scheme(font_scheme_elem, indent="  ")

        # Write back
        root_elem.write(theme_path)

    if theme_info.preserve_mono:
        logger.info("Preserving the existing monospace fonts as-is.")
    else:
        logger.info("Target monospace font:")
        logger.info("  latin: %s", theme_info.mono_font_latin)
        logger.info("  hangul: %s", theme_info.mono_font_hangul)


def _get_font_theme_dir() -> Path:
    match sys.platform:
        case "darwin":
            return (
                Path.home()
                / "Library"
                / "Group Containers"
                / "UBF8T346G9.Office"
                / "User Content.localized"
                / "Themes.localized"
                / "Theme Fonts"
            )
        case "win32":
            return Path.home() / "AppData" / "Roaming" / "Templates" / "Document Themes" / "Theme Fonts"
        case _:
            raise RuntimeError("Unsupported OS to auto-detect Microsoft Office's theme directory")


class FontThemeError(RuntimeError):
    """Raised when an Office font theme cannot be installed, with the message, the path and the reason as args."""

    def __str__(self) -> str:
        message, *details = (str(arg) for arg in self.args if arg is not None)
        return f"{message} ({', '.join(details)})" if details else message


class FontThemeExistsError(FontThemeError, FileExistsError):
    """Raised when installing an Office font theme would overwrite an existing one."""

    def __init__(self, path: Path) -> None:
        super().__init__("The target theme file already exist.", str(path))
        self.path = path


class InvalidFontThemeNameError(ValueError):
    """Raised when an Office font theme name is not usable as a file name."""


_invalid_font_theme_name_chars: Final = frozenset('/\\:*?"<>|')


def validate_font_theme_name(name: str) -> str:
    """Check that the font theme name is a safe file name on every platform and return it stripped."""
    name = name.strip()
    if not name:
        raise InvalidFontThemeNameError("The font theme name must not be empty.")
    if len(name) > 100:
        raise InvalidFontThemeNameError("The font theme name must be at most 100 characters long.")
    if any(c in _invalid_font_theme_name_chars or ord(c) < 0x20 or ord(c) == 0x7F for c in name):
        raise InvalidFontThemeNameError(
            'The font theme name must not contain control characters or any of / \\ : * ? " < > |.'
        )
    if name.endswith("."):
        raise InvalidFontThemeNameError("The font theme name must not end with a dot.")
    return name


def build_font_theme_xml(theme_info: Theme, theme_name: str) -> bytes:
    """Build the content of an Office font theme definition file."""
    root_elem = etree.Element(
        etree.QName(xmlns["a"], "fontScheme"),
        nsmap={k: v for k, v in xmlns.items() if k == "a"},  # filter only the "a" (drawingml) namespace
    )
    _fill_font_scheme(root_elem, theme_info)
    root_elem.set("name", theme_name)
    return etree.tostring(root_elem, pretty_print=True)


def install_font_theme(
    xml: bytes,
    theme_name: str,
    *,
    overwrite: bool = False,
    theme_dir: Path | None = None,
) -> Path:
    """Write an Office font theme definition into the Office theme directory and return its path."""
    theme_name = validate_font_theme_name(theme_name)
    if theme_dir is None:
        try:
            theme_dir = _get_font_theme_dir()
        except RuntimeError as e:
            raise FontThemeError(*e.args) from e
    if not theme_dir.is_dir():
        raise FontThemeError("The office theme directory does not exist.", str(theme_dir))
    theme_path = theme_dir / f"{theme_name}.xml"
    if theme_path.exists() and not overwrite:
        raise FontThemeExistsError(theme_path)
    try:
        theme_path.write_bytes(xml)
    except OSError as e:
        raise FontThemeError("Failed to write the theme file.", str(theme_path), e.strerror) from e
    return theme_path


def generate_font_theme(theme_info: Theme, theme_name: str, *, overwrite: bool = False) -> Path:
    theme_name = validate_font_theme_name(theme_name)
    theme_path = install_font_theme(build_font_theme_xml(theme_info, theme_name), theme_name, overwrite=overwrite)
    logger.info("Stored an Office theme font definition at:\n%s", theme_path)
    return theme_path


class FontName(NamedTuple):
    family: str
    weight: str | None = None
    """The weight variant such as "ExtraBold", or None for the regular weight."""
    slope: str | None = None
    """The slope variant, "Italic" or "Oblique", or None for the upright style."""

    @property
    def has_variant(self) -> bool:
        return self.weight is not None or self.slope is not None


def _split_font_name(typeface: str | None) -> FontName:
    """Split a typeface name such as "Pretendard ExtraBold Italic" into its family, weight and slope."""
    if not typeface:
        return FontName("")
    m = _variant_suffix_pattern.search(typeface)
    assert m is not None, "the pattern matches at least the end of the string"
    weight = m.group("weight")
    if weight is not None:
        if weight.lower() in ("regular", "normal"):
            weight = None
        else:
            prefix = m.group("prefix")
            weight = (prefix.capitalize() if prefix else "") + weight.capitalize()
    slope = m.group("slope")
    if slope is not None:
        slope = slope.capitalize()
    return FontName(typeface[: m.start()], weight, slope)


def _with_variant(typeface: str, variant: FontName) -> str:
    """
    Apply the weight and slope of the given variant to the typeface name, replacing its own ones.

    The typeface name is kept as-is when the variant has neither a weight nor a slope.
    """
    if not variant.has_variant:
        return typeface
    font_name = _split_font_name(typeface)
    weight = variant.weight or font_name.weight
    slope = variant.slope or font_name.slope
    return " ".join(part for part in (font_name.family, weight, slope) if part)


def _match_monospace_font(typeface: str | None) -> bool:
    if not typeface:
        return False
    return _split_font_name(typeface).family.lower() in known_monospace_fonts


def _has_monospace_font(prop_elem: etree._Element) -> bool:
    """Check if any typeface of the given text property element is a monospace font."""
    for elem in prop_elem:
        if local_tag(elem) not in preservable_typeface_tags:
            continue
        if _match_monospace_font(elem.get("typeface")):
            return True
    return False


def _has_font_variant(prop_elem: etree._Element) -> bool:
    """Check if any typeface of the given text property element has a non-regular weight or a slope variant."""
    for elem in prop_elem:
        if local_tag(elem) not in preservable_typeface_tags:
            continue
        if _split_font_name(elem.get("typeface")).has_variant:
            return True
    return False


def _update_paragraph_style(prop_elem: etree._Element, theme_info: Theme, scheme_prefix: str = "mn") -> None:
    if theme_info.preserve_mono and _has_monospace_font(prop_elem):
        return
    # Snapshot the children: the "font" case removes from the element being iterated.
    for elem in list(prop_elem):
        elem_name = local_tag(elem)
        typeface = elem.get("typeface")
        variant = _split_font_name(typeface)
        match elem_name:
            case "latin" | "ea" | "cs":
                if _match_monospace_font(typeface):
                    elem.clear()
                    if elem_name == "latin":
                        elem.set("typeface", _with_variant(theme_info.mono_font_latin, variant))
                    else:
                        elem.set("typeface", _with_variant(theme_info.mono_font_hangul, variant))
                elif variant.has_variant:
                    # The theme font reference cannot carry a weight or a slope, so name the theme font explicitly.
                    elem.clear()
                    match scheme_prefix, elem_name:
                        case "mn", "latin":
                            family = theme_info.minor_font_latin
                        case "mn", _:
                            family = theme_info.minor_font_hangul
                        case _, "latin":
                            family = theme_info.major_font_latin
                        case _, _:
                            family = theme_info.major_font_hangul
                    elem.set("typeface", _with_variant(family, variant))
                else:
                    elem.clear()
                    elem.set("typeface", f"+{scheme_prefix}-{style_typeface_map[elem_name]}")
            case "sym":
                elem.clear()
                elem.set(
                    "typeface",
                    _with_variant(
                        theme_info.minor_font_symbol if scheme_prefix == "mn" else theme_info.major_font_symbol,
                        variant,
                    ),
                )
            case "font":
                prop_elem.remove(elem)
            case _:
                pass


def _update_first_level_bullet_style(prop_elem: etree._Element, theme_info: Theme) -> None:
    first_level_style = theme_info.body_first_level_style
    assert first_level_style is not None, "the theme must define body_first_level_style"
    if theme_info.preserve_mono and _has_monospace_font(prop_elem):
        return
    for elem in prop_elem:
        elem_name = local_tag(elem)
        typeface = elem.get("typeface")
        # An explicit weight variant takes precedence over the theme's first-level style.
        variant = _split_font_name(typeface)
        variant = variant._replace(weight=variant.weight or first_level_style)
        match elem_name:
            case "latin":
                if _match_monospace_font(typeface):
                    elem.clear()
                    elem.set("typeface", _with_variant(theme_info.mono_font_latin, variant))
                else:
                    elem.clear()
                    elem.set("typeface", _with_variant(theme_info.minor_font_latin, variant))
            case "ea" | "cs":
                if _match_monospace_font(typeface):
                    elem.clear()
                    elem.set("typeface", _with_variant(theme_info.mono_font_hangul, variant))
                else:
                    elem.clear()
                    elem.set("typeface", _with_variant(theme_info.minor_font_hangul, variant))
            case _:
                pass


def normalize_master_fonts(
    work_path: Path,
    theme_info: Theme,
) -> None:
    main_path = work_path / "ppt" / "presentation.xml"
    root_elem = etree.parse(main_path)
    for prop_elem in xpath_elements(root_elem, "//p:defaultTextStyle//a:defRPr"):
        _update_paragraph_style(prop_elem, theme_info)

    master_dir = work_path / "ppt" / "slideMasters"
    for master_path in master_dir.glob("slideMaster*.xml"):
        root_elem = etree.parse(master_path)
        for prop_elem in xpath_elements(root_elem, "//p:titleStyle//a:defRPr"):
            _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mj")
            if theme_info.title_bold:
                prop_elem.set("b", "1")
            elif "b" in prop_elem.attrib:
                del prop_elem.attrib["b"]
        for prop_elem in xpath_elements(root_elem, "//p:bodyStyle//a:defRPr"):
            _update_paragraph_style(prop_elem, theme_info)
        for prop_elem in xpath_elements(root_elem, "//p:otherStyle//a:defRPr"):
            _update_paragraph_style(prop_elem, theme_info)
        if theme_info.body_first_level_style is not None:
            for prop_elem in xpath_elements(root_elem, "//p:bodyStyle//a:lvl1pPr//a:defRPr"):
                _update_first_level_bullet_style(prop_elem, theme_info)
        for bullet_font_elem in xpath_elements(root_elem, "//p:bodyStyle//a:buFont"):
            if theme_info.preserve_mono and _match_monospace_font(bullet_font_elem.get("typeface")):
                continue
            variant = _split_font_name(bullet_font_elem.get("typeface"))
            bullet_font_elem.clear()
            bullet_font_elem.set("typeface", _with_variant(theme_info.minor_font_symbol, variant))

        root_elem.write(master_path)


def _normalize_slide_font(root_elem: etree._ElementTree, theme_info: Theme, log_prefix: str) -> None:
    for sp_elem in xpath_elements(root_elem, "//p:sp"):
        if ph_elems := xpath_elements(sp_elem, "p:nvSpPr//p:ph"):
            match ph_elems[0].get("type", "body"):
                case "title":
                    scheme_prefix = "mj"
                case _:  # other values may be "body", "sldNum", ...
                    scheme_prefix = "mn"
            logger.info("%s: template element (%s)", log_prefix, ph_elems[0].get("type"))
            for prop_elem in xpath_elements(sp_elem, "p:txBody//a:defRPr"):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix=scheme_prefix)
            for prop_elem in xpath_elements(sp_elem, "p:txBody//a:rPr"):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix=scheme_prefix)
            for prop_elem in xpath_elements(sp_elem, "p:txBody//a:endParaRPr"):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix=scheme_prefix)
            if theme_info.body_first_level_style is not None:
                for prop_elem in xpath_elements(sp_elem, "//a:lvl1pPr//a:defRPr"):
                    _update_first_level_bullet_style(prop_elem, theme_info)
        else:
            logger.info("%s: normal element", log_prefix)
            for prop_elem in xpath_elements(sp_elem, "p:txBody//a:rPr"):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mn")
            for prop_elem in xpath_elements(sp_elem, "p:txBody//a:endParaRPr"):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mn")

    for tbl_elem in xpath_elements(root_elem, "//p:graphicFrame//a:tbl"):
        logger.info("%s: table element", log_prefix)
        for prop_elem in xpath_elements(tbl_elem, ".//a:txBody//a:defRPr"):
            _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mn")
        for prop_elem in xpath_elements(tbl_elem, ".//a:txBody//a:rPr"):
            _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mn")
        for prop_elem in xpath_elements(tbl_elem, ".//a:txBody//a:endParaRPr"):
            _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mn")
        for tx_style_elem in xpath_elements(tbl_elem, ".//a:tcTxStyle"):
            _update_table_text_style(tx_style_elem, theme_info)

    for bullet_font_elem in xpath_elements(root_elem, "//a:pPr//a:buFont"):
        if theme_info.preserve_mono and _match_monospace_font(bullet_font_elem.get("typeface")):
            continue
        variant = _split_font_name(bullet_font_elem.get("typeface"))
        bullet_font_elem.clear()
        bullet_font_elem.set("typeface", _with_variant(theme_info.minor_font_symbol, variant))


def normalize_layout_fonts(
    work_path: Path,
    theme_info: Theme,
) -> None:
    layout_dir = work_path / "ppt" / "slideLayouts"
    for layout_path in layout_dir.glob("slideLayout*.xml"):
        root_elem = etree.parse(layout_path)
        _normalize_slide_font(root_elem, theme_info, log_prefix=layout_path.name)
        root_elem.write(layout_path)


def normalize_slide_fonts(
    work_path: Path,
    theme_info: Theme,
) -> None:
    slide_dir = work_path / "ppt" / "slides"
    for slide_path in slide_dir.glob("slide*.xml"):
        root_elem = etree.parse(slide_path)
        _normalize_slide_font(root_elem, theme_info, log_prefix=slide_path.name)
        root_elem.write(slide_path)


def _update_table_text_style(tx_style_elem: etree._Element, theme_info: Theme) -> None:
    for font_elem in xpath_elements(tx_style_elem, "a:font"):
        if _has_monospace_font(font_elem) or _has_font_variant(font_elem):
            _update_paragraph_style(font_elem, theme_info)
            continue
        # Defer to the theme's minor font like the built-in table styles do.
        font_ref_elem = etree.Element(etree.QName(xmlns["a"], "fontRef"), attrib={"idx": "minor"})
        tx_style_elem.replace(font_elem, font_ref_elem)


def normalize_table_style_fonts(
    work_path: Path,
    theme_info: Theme,
) -> None:
    table_styles_path = work_path / "ppt" / "tableStyles.xml"
    if not table_styles_path.is_file():
        return
    root_elem = etree.parse(table_styles_path)
    for tx_style_elem in xpath_elements(root_elem, "//a:tcTxStyle"):
        _update_table_text_style(tx_style_elem, theme_info)
    root_elem.write(table_styles_path)


def _normalize_text_part_font(part_path: Path, theme_info: Theme) -> None:
    root_elem = etree.parse(part_path)
    for prop_elem in xpath_elements(root_elem, "//a:defRPr | //a:rPr | //a:endParaRPr"):
        _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mn")
    root_elem.write(part_path)


def normalize_chart_fonts(
    work_path: Path,
    theme_info: Theme,
) -> None:
    chart_dir = work_path / "ppt" / "charts"
    # Also matches the chartEx*.xml parts of the newer chart types.
    for chart_path in chart_dir.glob("chart*.xml"):
        logger.info("%s: chart", chart_path.name)
        _normalize_text_part_font(chart_path, theme_info)


def normalize_diagram_fonts(
    work_path: Path,
    theme_info: Theme,
) -> None:
    diagram_dir = work_path / "ppt" / "diagrams"
    # SmartArt keeps its text in the data model and a cached drawing that PowerPoint renders.
    for diagram_path in [*diagram_dir.glob("data*.xml"), *diagram_dir.glob("drawing*.xml")]:
        logger.info("%s: diagram", diagram_path.name)
        _normalize_text_part_font(diagram_path, theme_info)
