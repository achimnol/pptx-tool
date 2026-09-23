"""Generate a minimal pptx package that exercises every stage of the font-fixing pipeline.

It is not a complete presentation that PowerPoint can open, but it has all the parts the
pipeline reads and writes, so the tests do not need to commit any binary fixture.
"""

import zipfile
from pathlib import Path

from pptx_tool.fix import xmlns

_NSDECL = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in xmlns.items())


def _content_types() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        "</Types>"
    )


def _presentation(latin: str) -> str:
    return (
        f"<p:presentation {_NSDECL}><p:defaultTextStyle><a:lvl1pPr><a:defRPr>"
        f'<a:latin typeface="{latin}"/><a:ea typeface="{latin}"/>'
        "</a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>"
    )


def _theme(latin: str) -> str:
    fonts = (
        f'<a:latin typeface="{latin}"/><a:ea typeface=""/><a:cs typeface=""/>'
        '<a:font script="Hang" typeface="맑은 고딕"/>'
    )
    return (
        f'<a:theme {_NSDECL} name="Office Theme"><a:themeElements>'
        f'<a:fontScheme name="Office"><a:majorFont>{fonts}</a:majorFont><a:minorFont>{fonts}</a:minorFont>'
        "</a:fontScheme></a:themeElements></a:theme>"
    )


def _slide_master(latin: str) -> str:
    return (
        f"<p:sldMaster {_NSDECL}><p:txStyles>"
        f'<p:titleStyle><a:lvl1pPr><a:defRPr><a:latin typeface="{latin}"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
        "<p:bodyStyle><a:lvl1pPr>"
        f'<a:buFont typeface="Wingdings"/><a:defRPr><a:latin typeface="{latin}"/></a:defRPr>'
        "</a:lvl1pPr></p:bodyStyle>"
        f'<p:otherStyle><a:lvl1pPr><a:defRPr><a:latin typeface="{latin}"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
        "</p:txStyles></p:sldMaster>"
    )


def _slide_layout(latin: str) -> str:
    return (
        f"<p:sldLayout {_NSDECL}><p:cSld><p:spTree>"
        '<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody>'
        f'<a:p><a:r><a:rPr><a:latin typeface="{latin}"/></a:rPr><a:t>Title</a:t></a:r></a:p>'
        "</p:txBody></p:sp>"
        "</p:spTree></p:cSld></p:sldLayout>"
    )


def _slide(latin: str, mono: str) -> str:
    return (
        f"<p:sld {_NSDECL}><p:cSld><p:spTree>"
        "<p:sp><p:nvSpPr><p:nvPr/></p:nvSpPr><p:txBody><a:p>"
        f'<a:r><a:rPr><a:latin typeface="{latin}"/></a:rPr><a:t>Text</a:t></a:r>'
        f'<a:r><a:rPr><a:latin typeface="{mono}"/></a:rPr><a:t>code()</a:t></a:r>'
        "</a:p></p:txBody></p:sp>"
        "</p:spTree></p:cSld></p:sld>"
    )


def make_minimal_pptx(path: Path, *, latin: str = "Arial", mono: str = "Consolas") -> Path:
    """Write a minimal pptx package at the given path and return the path."""
    parts = {
        "[Content_Types].xml": _content_types(),
        "ppt/presentation.xml": _presentation(latin),
        "ppt/theme/theme1.xml": _theme(latin),
        "ppt/slideMasters/slideMaster1.xml": _slide_master(latin),
        "ppt/slideLayouts/slideLayout1.xml": _slide_layout(latin),
        "ppt/slides/slide1.xml": _slide(latin, mono),
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in parts.items():
            zf.writestr(name, content)
    return path
