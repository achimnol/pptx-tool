import logging
import sys
from pathlib import Path
from typing import Final, cast

from lxml import etree

from .types import Theme

logger = logging.getLogger(__name__)

xmlns: Final = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
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
    """Raised when an Office font theme cannot be installed."""


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


def _match_monospace_font(typeface: str | None) -> bool:
    if not typeface:
        return False
    return typeface.lower() in known_monospace_fonts


def _has_monospace_font(prop_elem: etree._Element) -> bool:
    """Check if any typeface of the given text property element is a monospace font."""
    for elem in prop_elem:
        if local_tag(elem) not in preservable_typeface_tags:
            continue
        if _match_monospace_font(elem.get("typeface")):
            return True
    return False


def _update_paragraph_style(prop_elem: etree._Element, theme_info: Theme, scheme_prefix: str = "mn") -> None:
    if theme_info.preserve_mono and _has_monospace_font(prop_elem):
        return
    # Snapshot the children: the "font" case removes from the element being iterated.
    for elem in list(prop_elem):
        elem_name = local_tag(elem)
        match elem_name:
            case "latin" | "ea" | "cs":
                if _match_monospace_font(elem.get("typeface")):
                    elem.clear()
                    if elem_name == "latin":
                        elem.set("typeface", theme_info.mono_font_latin)
                    else:
                        elem.set("typeface", theme_info.mono_font_hangul)
                else:
                    elem.clear()
                    elem.set("typeface", f"+{scheme_prefix}-{style_typeface_map[elem_name]}")
            case "sym":
                elem.clear()
                elem.set(
                    "typeface", theme_info.minor_font_symbol if scheme_prefix == "mn" else theme_info.major_font_symbol
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
        match elem_name:
            case "latin":
                if _match_monospace_font(elem.get("typeface")):
                    elem.clear()
                    elem.set("typeface", theme_info.mono_font_latin + " " + first_level_style)
                else:
                    elem.clear()
                    elem.set("typeface", theme_info.minor_font_latin + " " + first_level_style)
            case "ea" | "cs":
                if _match_monospace_font(elem.get("typeface")):
                    elem.clear()
                    elem.set("typeface", theme_info.mono_font_hangul + " " + first_level_style)
                else:
                    elem.clear()
                    elem.set("typeface", theme_info.minor_font_hangul + " " + first_level_style)
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
            bullet_font_elem.clear()
            bullet_font_elem.set("typeface", theme_info.minor_font_symbol)

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

    for bullet_font_elem in xpath_elements(root_elem, "//a:pPr//a:buFont"):
        if theme_info.preserve_mono and _match_monospace_font(bullet_font_elem.get("typeface")):
            continue
        bullet_font_elem.clear()
        bullet_font_elem.set("typeface", theme_info.minor_font_symbol)


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
