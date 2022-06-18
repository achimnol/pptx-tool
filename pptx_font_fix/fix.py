from typing import Final
from pathlib import Path

from lxml import etree

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


def fix_theme_font(
    work_path: Path,
    *,
    major_font: str,
    major_font_ko: str,
    minor_font: str,
    minor_font_ko: str,
) -> None:
    theme_dir = work_path / 'ppt' / 'theme'
    for theme_path in theme_dir.glob('theme*.xml'):
        root_elem = etree.parse(theme_path)
        font_scheme_elem = root_elem.xpath('//a:fontScheme', namespaces=xmlns)[0]

        # Print out current theme font configuration
        print(f"Current font scheme: (name={font_scheme_elem.get('name')!r})")
        _print_font_scheme(font_scheme_elem, indent="  ")

        # Replace the theme font
        major_font_elem = etree.Element(etree.QName(xmlns['a'], 'majorFont'))
        major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'latin'), attrib={'typeface': major_font}))
        major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'ea'), attrib={'typeface': major_font_ko}))
        major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'cs'), attrib={'typeface': major_font}))
        major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'sym'), attrib={'typeface': major_font}))
        major_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'font'), attrib={'script': 'Hang', 'typeface': major_font_ko}))
        minor_font_elem = etree.Element(etree.QName(xmlns['a'], 'minorFont'))
        minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'latin'), attrib={'typeface': minor_font}))
        minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'ea'), attrib={'typeface': minor_font_ko}))
        minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'cs'), attrib={'typeface': minor_font}))
        minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'sym'), attrib={'typeface': minor_font}))
        minor_font_elem.append(etree.Element(etree.QName(xmlns['a'], 'font'), attrib={'script': 'Hang', 'typeface': minor_font_ko}))
        font_scheme_attrib = {**font_scheme_elem.attrib}
        font_scheme_elem.clear()
        for k, v in font_scheme_attrib.items():
            font_scheme_elem.set(k, v)
        font_scheme_elem.append(major_font_elem)
        font_scheme_elem.append(minor_font_elem)

        print(f"New font scheme: (name={font_scheme_elem.get('name')!r})")
        _print_font_scheme(font_scheme_elem, indent="  ")

        # Write back
        root_elem.write(theme_path)


def _update_paragraph_style(style_elem: etree.Element, scheme_prefix: str = "mn") -> None:
    for elem in style_elem.getchildren():
        elem_name = local_tag(elem)
        match elem_name:
            case "latin" | "ea" | "cs" | "sym":
                elem.clear()
                elem.set('typeface', f"+{scheme_prefix}-{style_typeface_map[elem_name]}")
            case "font":
                elem.getparent().remove(elem)
            case _:
                pass


def normalize_master_fonts(
    work_path: Path,
    bullet_font: str = "+mn-lt",
    *,
    title_bold: bool = True,
    body_bold: bool = False,
) -> None:
    main_path = work_path / 'ppt' / 'presentation.xml'
    root_elem = etree.parse(main_path)
    for style_prop_elem in root_elem.xpath('//p:defaultTextStyle//a:defRPr', namespaces=xmlns):
        _update_paragraph_style(style_prop_elem)

    master_dir = work_path / 'ppt' / 'slideMasters'
    for master_path in master_dir.glob('slideMaster*.xml'):
        root_elem = etree.parse(master_path)
        for style_prop_elem in root_elem.xpath('//p:titleStyle//a:defRPr', namespaces=xmlns):
            _update_paragraph_style(style_prop_elem, scheme_prefix="mj")
            if title_bold:
                style_prop_elem.set('b', '1')
            else:
                style_prop_elem.attrib.pop('b', None)
        for style_prop_elem in root_elem.xpath('//p:bodyStyle//a:defRPr', namespaces=xmlns):
            _update_paragraph_style(style_prop_elem)
            if body_bold:
                style_prop_elem.set('b', '1')
            else:
                style_prop_elem.attrib.pop('b', None)
        for bullet_font_elem in root_elem.xpath('//p:bodyStyle//a:buFont', namespaces=xmlns):
            bullet_font_elem.clear()
            bullet_font_elem.set('typeface', bullet_font)

        root_elem.write(master_path)


def _normalize_slide_font(root_elem: etree.ElementTree, log_prefix: str) -> None:
    for sp_elem in root_elem.xpath('//p:sp', namespaces=xmlns):
        if ph_elems := sp_elem.xpath('p:nvSpPr//p:ph', namespaces=xmlns):
            match ph_elems[0].get('type', 'body'):
                case "title":
                    scheme_prefix = "mj"
                case _:  # other values may be "body", "sldNum", ...
                    scheme_prefix = "mn"
            print(f"{log_prefix}: template element ({ph_elems[0].get('type')})")
            for style_prop_elem in sp_elem.xpath('p:txBody//a:defRPr', namespaces=xmlns):
                _update_paragraph_style(style_prop_elem, scheme_prefix=scheme_prefix)
            for style_prop_elem in sp_elem.xpath('p:txBody//a:rPr', namespaces=xmlns):
                _update_paragraph_style(style_prop_elem, scheme_prefix=scheme_prefix)
            for style_prop_elem in sp_elem.xpath('p:txBody//a:endParaRPr', namespaces=xmlns):
                _update_paragraph_style(style_prop_elem, scheme_prefix=scheme_prefix)
        else:
            print(f"{log_prefix}: normal element")
            for style_prop_elem in sp_elem.xpath('p:txBody//a:rPr', namespaces=xmlns):
                _update_paragraph_style(style_prop_elem, scheme_prefix="mn")
            for style_prop_elem in sp_elem.xpath('p:txBody//a:endParaRPr', namespaces=xmlns):
                _update_paragraph_style(style_prop_elem, scheme_prefix="mn")



def normalize_layout_fonts(
    work_path: Path,
) -> None:
    layout_dir = work_path / 'ppt' / 'slideLayouts'
    for layout_path in layout_dir.glob('slideLayout*.xml'):
        root_elem = etree.parse(layout_path)
        _normalize_slide_font(root_elem, log_prefix=layout_path.name)
        root_elem.write(layout_path)


def normalize_slide_fonts(work_path: Path) -> None:
    slide_dir = work_path / 'ppt' / 'slides'
    for slide_path in slide_dir.glob('slide*.xml'):
        root_elem = etree.parse(slide_path)
        _normalize_slide_font(root_elem, log_prefix=slide_path.name)
        root_elem.write(slide_path)