import argparse
import json
import tempfile
from pathlib import Path

from .fix import (
    fix_theme_font,
    generate_font_theme,
    normalize_master_fonts,
    normalize_layout_fonts,
    normalize_slide_fonts,
)
from .package import extract_pptx, build_pptx
from .types import Theme


def _load_theme(args: argparse.Namespace) -> Theme:
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
    return theme_info

def do_fix_pptx(args: argparse.Namespace) -> None:
    theme_info = _load_theme(args)
    with tempfile.TemporaryDirectory(prefix="pptx-font-fix-") as tmp_dir:
        tmp_path = Path(tmp_dir)
        extract_pptx(args.src, tmp_path)
        fix_theme_font(tmp_path, theme_info)
        normalize_master_fonts(tmp_path, theme_info)
        normalize_layout_fonts(tmp_path, theme_info)
        normalize_slide_fonts(tmp_path, theme_info)
        build_pptx(tmp_path, args.dst)


def do_generate_font_theme(args: argparse.Namespace) -> None:
    theme_info = _load_theme(args)
    generate_font_theme(theme_info, args.name, overwrite=args.overwrite)


def main() -> None:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(title="commands")

    parser_fix = subparsers.add_parser(
        'fix-font',
        help="Fix up the font theme and normalize all slide objects to use major/minor fonts correctly in a pptx file.",
    )
    parser_fix.add_argument('--theme', type=Path, required=True, help="The path to a theme json file.")
    parser_fix.add_argument('src', type=Path, help="The source pptx file.")
    parser_fix.add_argument('dst', type=Path, help="The destination pptx file. You may set it same to `src`.")
    parser_fix.set_defaults(func=do_fix_pptx)

    parser_gen = subparsers.add_parser(
        'generate-font-theme',
        help="Generate and register an Office font theme. "
             "Note that the font theme only defines major/minor typeface and "
             "things like making slide titles bold should be done with slide master templates. "
             "It also does not support 'body-first-line-style' option. "
             "To use the new font theme, you must restart Office apps to take effect.",
    )
    parser_gen.add_argument('--theme', type=Path, required=True, help="The path to a theme json file.")
    parser_gen.add_argument('--overwrite', action='store_true', default=False, help="Overwrite the Office font theme file if already exists")
    parser_gen.add_argument('name', type=str, help="The name for your Office font theme. It must also be a valid file name.")
    parser_gen.set_defaults(func=do_generate_font_theme)

    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
