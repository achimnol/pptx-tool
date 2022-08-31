import sys
from pathlib import Path
from typing import Final

from lxml import etree

from .types import Theme

xmlns: Final = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

style_typeface_map: Final = {
    'latin': 'lt',
    'ea': 'ea',
    'cs': 'cs',
    'sym': 'sym',
}

known_monospace_fonts: Final = {
    'courier',
    'courier new',
    'consolas',
    'cascadia code',
    'cascadia mono',
    'inconsolata',
    'jetbrains mono',
    'menlo',
    'monaco',
    'terminus',
    'bitstream vera sans mono',
    'droid sans mono',
    'dejavu sans mono',
    'pt mono',
    'sf mono',
    'andale mono',
    'arpercu mono',
    'dank mono',
    'input mono',
    'roboto mono',
    'oxygen mono',
    'space mono',
    'ubuntu mono',
    'liberation mono',
    'anonymous pro',
    'source code pro',
    'iosevka',
    'monolisa',
    'monoid',
    'gintronic',
    'fira code',
    'nanumgothiccoding',
    'hack',
    'recursive',
}


def local_tag(tag_name) -> str:
    """Strip out the namespace from the element tag name."""
    return etree.QName(tag_name).localname


def _print_font_scheme(font_scheme: etree.Element, indent: str = "") -> None:
    for font_elem in font_scheme.getchildren():
        print(f"{indent}{local_tag(font_elem.tag)}:")
        for prev_typeface in font_elem.getchildren():
            prev_script_name = local_tag(prev_typeface.tag)
            match prev_script_name:
                case "font":
                    print(
                        f"{indent}  "
                        f"{prev_script_name} ({prev_typeface.get('script')}): "
                        f"{prev_typeface.get('typeface')}"
                    )
                case _:
                    print(
                        f"{indent}  "
                        f"{prev_script_name}: {prev_typeface.get('typeface')}"
                    )


def _fill_font_scheme(target_elem: etree.Element, theme_info: Theme) -> None:
    major_font_elem = etree.Element(etree.QName(xmlns['a'], 'majorFont'))
    major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'latin'), attrib={'typeface': theme_info.major_font_latin}))
    major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'ea'), attrib={'typeface': theme_info.major_font_hangul}))
    major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'cs'), attrib={'typeface': theme_info.major_font_hangul}))
    major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'sym'), attrib={'typeface': theme_info.major_font_symbol}))
    # major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'font'), attrib={'script': 'Hang', 'typeface': theme_info.major_font_hangul}))
    minor_font_elem = etree.Element(etree.QName(xmlns['a'], 'minorFont'))
    minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'latin'), attrib={'typeface': theme_info.minor_font_latin}))
    minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'ea'), attrib={'typeface': theme_info.minor_font_hangul}))
    minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'cs'), attrib={'typeface': theme_info.minor_font_hangul}))
    minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'sym'), attrib={'typeface': theme_info.minor_font_symbol}))
    # minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'font'), attrib={'script': 'Hang', 'typeface': theme_info.minor_font_hangul}))
    # Replace the theme font
    font_scheme_attrib = {**target_elem.attrib}
    target_elem.clear()
    for k, v in font_scheme_attrib.items():
        target_elem.set(k, v)
    target_elem.append(major_font_elem)
    target_elem.append(minor_font_elem)


def fix_theme_font(
    work_path: Path,
    theme_info: Theme,
) -> None:
    theme_dir = work_path / 'ppt' / 'theme'
    for theme_path in theme_dir.glob('theme*.xml'):
        root_elem = etree.parse(theme_path)
        font_scheme_elem = root_elem.xpath('//a:fontScheme', namespaces=xmlns)[0]

        # Print out current theme font configuration
        print(f"Current font scheme: (name={font_scheme_elem.get('name')!r})")
        _print_font_scheme(font_scheme_elem, indent="  ")

        _fill_font_scheme(font_scheme_elem, theme_info)

        print(f"New font scheme: (name={font_scheme_elem.get('name')!r})")
        _print_font_scheme(font_scheme_elem, indent="  ")

        # Write back
        root_elem.write(theme_path)

    print("Target monospace font:")
    print(f"  latin: {theme_info.mono_font_latin}")
    print(f"  hangul: {theme_info.mono_font_hangul}")


def _get_font_theme_dir() -> Path:
    match sys.platform:
        case 'darwin':
            return (
                Path.home() / 'Library' / 'Group Containers'
                / 'UBF8T346G9.Office' / 'User Content.localized'
                / 'Themes.localized' / 'Theme Fonts'
            )
        case 'win32':
            return (
                Path.home() / 'AppData' / 'Roaming'
                / 'Templates' / 'Document Themes' / 'Theme Fonts'
            )
        case _:
            raise RuntimeError("Unsupported OS to auto-detect Microsoft Office's theme directory")


def generate_font_theme(theme_info: Theme, theme_name: str, *, overwrite: bool = False) -> None:
    theme_dir = _get_font_theme_dir()
    if not theme_dir.is_dir():
        raise RuntimeError("The office theme directory does not exist.", str(theme_dir))
    theme_path = theme_dir / f"{theme_name}.xml"
    if theme_path.exists() and not overwrite:
        raise RuntimeError("The target theme file already exist.", str(theme_path))
    root_elem = etree.Element(
        etree.QName(xmlns['a'], 'fontScheme'),
        nsmap={k: v for k, v in xmlns.items() if k == "a"},  # filter only the "a" (drawingml) namespace
    )
    _fill_font_scheme(root_elem, theme_info)
    root_elem.set('name', theme_name)
    tree = etree.ElementTree(root_elem)
    tree.write(theme_path, pretty_print=True)
    print(f"Stored an Office theme font definition at:\n{theme_path}")


def _match_monospace_font(typeface: str) -> bool:
    return typeface.lower() in known_monospace_fonts


def _update_paragraph_style(prop_elem: etree.Element, theme_info: Theme, scheme_prefix: str = "mn") -> None:
    for elem in prop_elem.getchildren():
        elem_name = local_tag(elem)
        match elem_name:
            case "latin" | "ea" | "cs":
                if _match_monospace_font(elem.get('typeface')):
                    elem.clear()
                    if elem_name == "latin":
                        elem.set('typeface', theme_info.mono_font_latin)
                    else:
                        elem.set('typeface', theme_info.mono_font_hangul)
                else:
                    elem.clear()
                    elem.set('typeface', f"+{scheme_prefix}-{style_typeface_map[elem_name]}")
            case "sym":
                elem.clear()
                elem.set('typeface', theme_info.minor_font_symbol if scheme_prefix == "mn" else theme_info.major_font_symbol)
            case "font":
                elem.getparent().remove(elem)
            case _:
                pass


def _update_first_level_bullet_style(prop_elem: etree.Element, theme_info: Theme) -> None:
    for elem in prop_elem.getchildren():
        elem_name = local_tag(elem)
        match elem_name:
            case "latin":
                if _match_monospace_font(elem.get('typeface')):
                    elem.clear()
                    elem.set('typeface', theme_info.mono_font_latin + " " + theme_info.body_first_level_style)
                else:
                    elem.clear()
                    elem.set('typeface', theme_info.minor_font_latin + " " + theme_info.body_first_level_style)
            case "ea" | "cs":
                if _match_monospace_font(elem.get('typeface')):
                    elem.clear()
                    elem.set('typeface', theme_info.mono_font_hangul + " " + theme_info.body_first_level_style)
                else:
                    elem.clear()
                    elem.set('typeface', theme_info.minor_font_hangul + " " + theme_info.body_first_level_style)
            case _:
                pass


def normalize_master_fonts(
    work_path: Path,
    theme_info: Theme,
) -> None:
    main_path = work_path / 'ppt' / 'presentation.xml'
    root_elem = etree.parse(main_path)
    for prop_elem in root_elem.xpath('//p:defaultTextStyle//a:defRPr', namespaces=xmlns):
        _update_paragraph_style(prop_elem, theme_info)

    master_dir = work_path / 'ppt' / 'slideMasters'
    for master_path in master_dir.glob('slideMaster*.xml'):
        root_elem = etree.parse(master_path)
        for prop_elem in root_elem.xpath('//p:titleStyle//a:defRPr', namespaces=xmlns):
            _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mj")
            if theme_info.title_bold:
                prop_elem.set('b', '1')
            else:
                prop_elem.attrib.pop('b', None)
        for prop_elem in root_elem.xpath('//p:bodyStyle//a:defRPr', namespaces=xmlns):
            _update_paragraph_style(prop_elem, theme_info)
        for prop_elem in root_elem.xpath('//p:otherStyle//a:defRPr', namespaces=xmlns):
            _update_paragraph_style(prop_elem, theme_info)
        if theme_info.body_first_level_style is not None:
            for prop_elem in root_elem.xpath('//p:bodyStyle//a:lvl1pPr//a:defRPr', namespaces=xmlns):
                _update_first_level_bullet_style(prop_elem, theme_info)
        for bullet_font_elem in root_elem.xpath('//p:bodyStyle//a:buFont', namespaces=xmlns):
            bullet_font_elem.clear()
            bullet_font_elem.set('typeface', theme_info.minor_font_symbol)

        root_elem.write(master_path)


def _normalize_slide_font(root_elem: etree.ElementTree, theme_info: Theme, log_prefix: str) -> None:
    for sp_elem in root_elem.xpath('//p:sp', namespaces=xmlns):
        if ph_elems := sp_elem.xpath('p:nvSpPr//p:ph', namespaces=xmlns):
            match ph_elems[0].get('type', 'body'):
                case "title":
                    scheme_prefix = "mj"
                case _:  # other values may be "body", "sldNum", ...
                    scheme_prefix = "mn"
            print(f"{log_prefix}: template element ({ph_elems[0].get('type')})")
            for prop_elem in sp_elem.xpath('p:txBody//a:defRPr', namespaces=xmlns):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix=scheme_prefix)
            for prop_elem in sp_elem.xpath('p:txBody//a:rPr', namespaces=xmlns):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix=scheme_prefix)
            for prop_elem in sp_elem.xpath('p:txBody//a:endParaRPr', namespaces=xmlns):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix=scheme_prefix)
            if theme_info.body_first_level_style is not None:
                for prop_elem in sp_elem.xpath('//a:lvl1pPr//a:defRPr', namespaces=xmlns):
                    _update_first_level_bullet_style(prop_elem, theme_info)
        else:
            print(f"{log_prefix}: normal element")
            for prop_elem in sp_elem.xpath('p:txBody//a:rPr', namespaces=xmlns):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mn")
            for prop_elem in sp_elem.xpath('p:txBody//a:endParaRPr', namespaces=xmlns):
                _update_paragraph_style(prop_elem, theme_info, scheme_prefix="mn")



def normalize_layout_fonts(
    work_path: Path,
    theme_info: Theme,
) -> None:
    layout_dir = work_path / 'ppt' / 'slideLayouts'
    for layout_path in layout_dir.glob('slideLayout*.xml'):
        root_elem = etree.parse(layout_path)
        _normalize_slide_font(root_elem, theme_info, log_prefix=layout_path.name)
        root_elem.write(layout_path)


def normalize_slide_fonts(
    work_path: Path,
    theme_info: Theme,
) -> None:
    slide_dir = work_path / 'ppt' / 'slides'
    for slide_path in slide_dir.glob('slide*.xml'):
        root_elem = etree.parse(slide_path)
        _normalize_slide_font(root_elem, theme_info, log_prefix=slide_path.name)
        root_elem.write(slide_path)
