import argparse
import contextlib
import dataclasses
import logging
import sys
from pathlib import Path

from .fix import FontThemeError, InvalidFontThemeNameError, generate_font_theme
from .log import cli_logging
from .package import build_pptx, extract_pptx
from .pipeline import fix_pptx
from .theme import ThemeError, resolve_theme_arg
from .types import Theme

_THEME_HELP = "The path to a theme json file, or the id of a bundled theme, such as 'pretendard'."


def _resolve_preserve_mono(theme_info: Theme, args: argparse.Namespace) -> Theme:
    """Let an explicit `--preserve-mono`/`--no-preserve-mono` override the theme option."""
    preserve_mono = getattr(args, "preserve_mono", None)
    if preserve_mono is None:
        return theme_info
    return dataclasses.replace(theme_info, preserve_mono=preserve_mono)


def do_extract_pptx(args: argparse.Namespace) -> None:
    dst_path: Path = args.dst
    if dst_path.exists():
        assert dst_path.is_dir()
    else:
        dst_path.mkdir(parents=True)
    extract_pptx(args.src, dst_path)


def do_build_pptx(args: argparse.Namespace) -> None:
    build_pptx(args.src, args.dst)


def do_fix_pptx(args: argparse.Namespace) -> None:
    theme_info = _resolve_preserve_mono(resolve_theme_arg(args.theme), args)
    fix_pptx(args.src, args.dst, theme_info)


def do_generate_font_theme(args: argparse.Namespace) -> None:
    theme_info = resolve_theme_arg(args.theme)
    generate_font_theme(theme_info, args.name, overwrite=args.overwrite)


_WEB_EXTRA_HINT = (
    "The web UI requires the 'web' extra. Install it with `uv sync --extra web` or `pip install 'pptx-tool[web]'`."
)


def do_serve(args: argparse.Namespace) -> None:
    try:
        import uvicorn

        from .web.app import create_app
        from .web.config import LOOPBACK_ADDRESSES, WebConfig, loopback_host_headers
    except ImportError as e:
        if e.name is not None and e.name.split(".")[0] in {"litestar", "uvicorn", "msgspec"}:
            sys.exit(_WEB_EXTRA_HINT)
        raise
    is_loopback = args.host in LOOPBACK_ADDRESSES
    if args.local and not is_loopback:
        sys.exit("The --local option is allowed only when serving on a loopback address such as 127.0.0.1.")
    if args.tmp_quota_mb < 2 * args.max_upload_mb:
        sys.exit("The --tmp-quota-mb option must be at least twice --max-upload-mb.")
    web_config = WebConfig(
        local=args.local,
        max_upload_size=args.max_upload_mb * 1024 * 1024,
        tmp_dir=args.tmp_dir,
        tmp_quota=args.tmp_quota_mb * 1024 * 1024,
        # Accept any Host header when serving on a public address, as the host names are unknown.
        allowed_hosts=loopback_host_headers(args.port) if is_loopback else (),
    )
    # Keep the processing logs of each request (with the user's font names) in its response only,
    # instead of also echoing them through the server's root log handler.
    logging.getLogger("pptx_tool").propagate = False
    uvicorn.run(create_app(web_config), host=args.host, port=args.port)


def main() -> None:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(title="commands")

    parser_extract = subparsers.add_parser(
        "extract",
        help="Extract the pptx file into a directory.",
    )
    parser_extract.add_argument("src", type=Path, help="The source pptx file.")
    parser_extract.add_argument("dst", type=Path, help="The destination directory to extract.")
    parser_extract.set_defaults(func=do_extract_pptx)

    parser_build = subparsers.add_parser(
        "build",
        help="Build the directory as a pptx file.",
    )
    parser_build.add_argument("src", type=Path, help="The source directory.")
    parser_build.add_argument("dst", type=Path, help="The destination pptx file.")
    parser_build.set_defaults(func=do_build_pptx)

    parser_fix = subparsers.add_parser(
        "fix-font",
        help="Fix up the font theme and normalize all slide objects to use major/minor fonts correctly in a pptx file.",
    )
    parser_fix.add_argument("--theme", required=True, help=_THEME_HELP)
    parser_fix.add_argument(
        "--preserve-mono",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Leave the text objects that already use a known monospace font untouched, "
        "instead of replacing them with the theme's monospace font. "
        "It overrides the theme's 'preserveMono' option when given. "
        "The default is the theme's option, which is off unless set.",
    )
    parser_fix.add_argument("src", type=Path, help="The source pptx file.")
    parser_fix.add_argument("dst", type=Path, help="The destination pptx file. You may set it same to `src`.")
    parser_fix.set_defaults(func=do_fix_pptx)

    parser_gen = subparsers.add_parser(
        "generate-font-theme",
        help="Generate and register an Office font theme. "
        "Note that the font theme only defines major/minor typeface and "
        "things like making slide titles bold should be done with slide master templates. "
        "It also does not support 'body-first-line-style' and 'preserveMono' options. "
        "To use the new font theme, you must restart Office apps to take effect.",
    )
    parser_gen.add_argument("--theme", required=True, help=_THEME_HELP)
    parser_gen.add_argument(
        "--overwrite", action="store_true", default=False, help="Overwrite the Office font theme file if already exists"
    )
    parser_gen.add_argument(
        "name", type=str, help="The name for your Office font theme. It must also be a valid file name."
    )
    parser_gen.set_defaults(func=do_generate_font_theme)

    parser_serve = subparsers.add_parser(
        "serve",
        help="Run the web UI server. It requires the 'web' extra.",
    )
    parser_serve.add_argument("--host", default="127.0.0.1", help="The address to listen on. (default: %(default)s)")
    parser_serve.add_argument("--port", type=int, default=8000, help="The port to listen on. (default: %(default)s)")
    parser_serve.add_argument(
        "--local",
        action="store_true",
        default=False,
        help="Enable the features that modify this machine, such as installing Office font themes. "
        "It is allowed only when listening on a loopback address.",
    )
    parser_serve.add_argument(
        "--max-upload-mb",
        type=int,
        default=200,
        help="The maximum size of uploaded pptx files in MiB. (default: %(default)s)",
    )
    parser_serve.add_argument(
        "--tmp-dir",
        type=Path,
        default=Path("/tmp/pptx-tool"),
        help="The directory for the uploaded and intermediate files, one subdirectory per request. "
        "(default: %(default)s)",
    )
    parser_serve.add_argument(
        "--tmp-quota-mb",
        type=int,
        default=1024,
        help="The total size limit of --tmp-dir in MiB, at least twice --max-upload-mb. "
        "Any leftover files of earlier requests are deleted to make room. (default: %(default)s)",
    )
    parser_serve.set_defaults(func=do_serve)

    args = parser.parse_args()
    try:
        # The server does not echo the processing logs of each request.
        with cli_logging() if args.func is not do_serve else contextlib.nullcontext():
            args.func(args)
    except (ThemeError, InvalidFontThemeNameError) as e:
        parser.error(str(e))
    except FontThemeError as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
