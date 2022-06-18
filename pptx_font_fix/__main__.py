import argparse
import json
import tempfile
from pathlib import Path

from .fix import (
    fix_theme_font,
    normalize_master_fonts,
    normalize_layout_fonts,
    normalize_slide_fonts,
)
from .package import extract_pptx, build_pptx
from .types import Theme


def do_fix_fonts(args: argparse.Namespace) -> None:
    theme_data = json.loads(args.theme.read_text())
    theme_info = Theme(
        major_font_latin=theme_data['majorFont']['latin'],
        major_font_hangul=theme_data['majorFont']['hangul'],
        major_font_symbol=theme_data['majorFont']['symbol'],
        minor_font_latin=theme_data['minorFont']['latin'],
        minor_font_hangul=theme_data['minorFont']['hangul'],
        minor_font_symbol=theme_data['minorFont']['symbol'],
        title_bold=theme_data['options']['titleBold'],
        body_first_level_style=theme_data['options']['bodyFirstLevelStyle'],
    )
    with tempfile.TemporaryDirectory(prefix="pptx-font-fix-") as tmp_dir:
        tmp_path = Path(tmp_dir)
        extract_pptx(args.src, tmp_path)
        fix_theme_font(tmp_path, theme_info)
        normalize_master_fonts(tmp_path, theme_info)
        normalize_layout_fonts(tmp_path, theme_info)
        normalize_slide_fonts(tmp_path, theme_info)
        build_pptx(tmp_path, args.dst)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument('--theme', type=Path)
    parser.add_argument('src', type=Path)
    parser.add_argument('dst', type=Path)
    args = parser.parse_args()
    do_fix_fonts(args)


if __name__ == '__main__':
    main()
