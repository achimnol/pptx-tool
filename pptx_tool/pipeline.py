import tempfile
from pathlib import Path

from .fix import (
    fix_theme_font,
    normalize_chart_fonts,
    normalize_diagram_fonts,
    normalize_layout_fonts,
    normalize_master_fonts,
    normalize_slide_fonts,
    normalize_table_style_fonts,
)
from .package import ArchiveLimits, build_pptx, extract_pptx
from .types import Theme


def fix_pptx(
    src: Path,
    dst: Path,
    theme_info: Theme,
    *,
    limits: ArchiveLimits | None = None,
    work_dir: Path | None = None,
) -> None:
    """
    Fix up the font theme and normalize all slide objects of the src pptx file into dst.

    The package is extracted into work_dir if given (which must be empty or not exist yet),
    or into a temporary directory otherwise.
    """
    if work_dir is not None:
        _fix_extracted_pptx(src, dst, theme_info, work_dir, limits)
        return
    with tempfile.TemporaryDirectory(prefix="pptx-font-fix-") as tmp_dir:
        _fix_extracted_pptx(src, dst, theme_info, Path(tmp_dir), limits)


def _fix_extracted_pptx(src: Path, dst: Path, theme_info: Theme, work_path: Path, limits: ArchiveLimits | None) -> None:
    extract_pptx(src, work_path, limits=limits)
    fix_theme_font(work_path, theme_info)
    normalize_master_fonts(work_path, theme_info)
    normalize_layout_fonts(work_path, theme_info)
    normalize_slide_fonts(work_path, theme_info)
    normalize_table_style_fonts(work_path, theme_info)
    normalize_chart_fonts(work_path, theme_info)
    normalize_diagram_fonts(work_path, theme_info)
    build_pptx(work_path, dst)
