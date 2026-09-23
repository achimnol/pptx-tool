import base64
import tempfile
import zipfile
from pathlib import Path, PurePath
from typing import Annotated, Any
from urllib.parse import quote, urlsplit

from litestar import Response, get, post
from litestar.connection import ASGIConnection
from litestar.di import NamedDependency
from litestar.enums import MediaType, RequestEncodingType
from litestar.exceptions import HTTPException, PermissionDeniedException, ValidationException
from litestar.handlers.base import BaseRouteHandler
from litestar.params import Body
from litestar.status_codes import (
    HTTP_201_CREATED,
    HTTP_409_CONFLICT,
    HTTP_413_REQUEST_ENTITY_TOO_LARGE,
    HTTP_422_UNPROCESSABLE_ENTITY,
    HTTP_503_SERVICE_UNAVAILABLE,
)
from lxml import etree

from ..fix import (
    FontThemeError,
    FontThemeExistsError,
    InvalidFontThemeNameError,
    build_font_theme_xml,
    install_font_theme,
    known_monospace_fonts,
    validate_font_theme_name,
)
from ..log import capture_log
from ..package import UnsafeArchiveError
from ..pipeline import fix_pptx
from ..theme import list_bundled_themes
from .config import WebConfig
from .models import (
    AppConfig,
    BundledThemeInfo,
    FixFontForm,
    FixFontResult,
    FontThemeRequest,
    InstallFontThemeRequest,
    InstallFontThemeResult,
    decode_theme_json,
    from_core_theme,
    to_core_theme,
)

_COPY_CHUNK_SIZE = 1024 * 1024


def same_origin_guard(connection: ASGIConnection[Any, Any, Any, Any], _: BaseRouteHandler) -> None:
    """Reject cross-site requests, which browsers mark with a foreign Origin header."""
    origin = connection.headers.get("origin")
    if origin is None:
        return
    if urlsplit(origin).netloc != connection.headers.get("host"):
        raise PermissionDeniedException("Cross-origin requests are not allowed.")


def _validate_font_theme_name(name: str) -> str:
    try:
        return validate_font_theme_name(name)
    except InvalidFontThemeNameError as e:
        raise ValidationException(detail=str(e), extra=[{"key": "name", "message": str(e), "source": "body"}]) from e


@get("/api/themes", sync_to_thread=False)
def list_themes() -> list[BundledThemeInfo]:
    """List the bundled themes."""
    return [BundledThemeInfo(t.id, t.name, from_core_theme(t.theme)) for t in list_bundled_themes()]


@get("/api/monospace-fonts", sync_to_thread=False)
def list_monospace_fonts() -> list[str]:
    """List the (lowercased) font names regarded as monospace fonts."""
    return sorted(known_monospace_fonts)


@get("/api/config", sync_to_thread=False)
def get_config(web_config: NamedDependency[WebConfig]) -> AppConfig:
    """Describe the server configuration that affects the UI."""
    return AppConfig(local=web_config.local, max_upload_size=web_config.max_upload_size)


@post("/api/fix-font", status_code=200, sync_to_thread=True)
def fix_font(
    data: Annotated[FixFontForm, Body(media_type=RequestEncodingType.MULTI_PART)],
    web_config: NamedDependency[WebConfig],
) -> FixFontResult:
    """Fix up the fonts of the uploaded pptx file with the given theme."""
    theme_info = decode_theme_json(data.theme)
    stem = PurePath(data.file.filename or "presentation").stem or "presentation"
    with tempfile.TemporaryDirectory(prefix="pptx-tool-web-") as tmp_dir:
        src_path = Path(tmp_dir) / "src.pptx"
        dst_path = Path(tmp_dir) / "dst.pptx"
        size = 0
        with src_path.open("wb") as f:
            while chunk := data.file.file.read(_COPY_CHUNK_SIZE):
                size += len(chunk)
                if size > web_config.max_upload_size:
                    raise HTTPException(
                        status_code=HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                        detail=f"The file is larger than the limit ({web_config.max_upload_size} bytes).",
                    )
                f.write(chunk)
        with capture_log() as log:
            try:
                fix_pptx(src_path, dst_path, theme_info, limits=web_config.archive_limits)
            except (zipfile.BadZipFile, UnsafeArchiveError) as e:
                raise ValidationException(detail=f"Not a valid pptx file: {e}") from e
            except (OSError, etree.XMLSyntaxError, IndexError) as e:
                raise HTTPException(
                    status_code=HTTP_422_UNPROCESSABLE_ENTITY,
                    detail=f"Failed to process the presentation: {e}",
                ) from e
        content = dst_path.read_bytes()
    return FixFontResult(
        filename=f"{stem}-fixed.pptx",
        log=log.text,
        content_base64=base64.b64encode(content).decode("ascii"),
    )


@post("/api/font-theme", status_code=200, sync_to_thread=False, media_type=MediaType.XML)
def font_theme(data: FontThemeRequest) -> Response[bytes]:
    """Generate an Office font theme definition file."""
    name = _validate_font_theme_name(data.name)
    xml = build_font_theme_xml(to_core_theme(data.theme), name)
    return Response(
        xml,
        media_type=MediaType.XML,
        headers={
            "Content-Disposition": f"attachment; filename=\"font-theme.xml\"; filename*=UTF-8''{quote(name)}.xml",
        },
    )


@post("/api/font-theme/install", status_code=HTTP_201_CREATED, sync_to_thread=True, guards=[same_origin_guard])
def install_font_theme_handler(
    data: InstallFontThemeRequest,
    web_config: NamedDependency[WebConfig],
) -> InstallFontThemeResult:
    """Install an Office font theme into the server machine. Available only in the local mode."""
    if not web_config.local:
        raise PermissionDeniedException("Installing font themes is available only when the server runs with --local.")
    name = _validate_font_theme_name(data.name)
    xml = build_font_theme_xml(to_core_theme(data.theme), name)
    try:
        path = install_font_theme(xml, name, overwrite=data.overwrite, theme_dir=web_config.font_theme_dir)
    except FontThemeExistsError as e:
        raise HTTPException(
            status_code=HTTP_409_CONFLICT,
            detail=f"A font theme named {name!r} already exists.",
            extra={"path": str(e.path)},
        ) from e
    except FontThemeError as e:
        raise HTTPException(
            status_code=HTTP_503_SERVICE_UNAVAILABLE,
            detail=" ".join(str(arg) for arg in e.args),
        ) from e
    return InstallFontThemeResult(path=str(path))


route_handlers = [
    list_themes,
    list_monospace_fonts,
    get_config,
    fix_font,
    font_theme,
    install_font_theme_handler,
]
