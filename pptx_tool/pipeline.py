import tempfile
from pathlib import Path

from .fix import fix_theme_font, normalize_layout_fonts, normalize_master_fonts, normalize_slide_fonts
from .package import ArchiveLimits, build_pptx, extract_pptx
from .types import Theme


def fix_pptx(src: Path, dst: Path, theme_info: Theme, *, limits: ArchiveLimits | None = None) -> None:
    """Fix up the font theme and normalize all slide objects of the src pptx file into dst."""
    with tempfile.TemporaryDirectory(prefix="pptx-font-fix-") as tmp_dir:
        tmp_path = Path(tmp_dir)
        extract_pptx(src, tmp_path, limits=limits)
        fix_theme_font(tmp_path, theme_info)
        normalize_master_fonts(tmp_path, theme_info)
        normalize_layout_fonts(tmp_path, theme_info)
        normalize_slide_fonts(tmp_path, theme_info)
        build_pptx(tmp_path, dst)
