from typing import Annotated, Any

import msgspec
from litestar.datastructures import UploadFile
from litestar.exceptions import ValidationException

from ..theme import ThemeError, dump_theme, load_theme
from ..types import Theme


class FontSet(msgspec.Struct):
    latin: str
    hangul: str
    symbol: str


class MonoFontSet(msgspec.Struct):
    latin: str
    hangul: str


class ThemeOptions(msgspec.Struct, rename="camel"):
    title_bold: bool
    body_first_level_style: str | None
    preserve_mono: bool = False


class ThemeData(msgspec.Struct, rename="camel"):
    """The theme definition, in the same structure as the theme JSON files."""

    major_font: FontSet
    minor_font: FontSet
    mono_font: MonoFontSet
    options: ThemeOptions


class BundledThemeInfo(msgspec.Struct):
    id: str
    name: str
    theme: ThemeData


class AppConfig(msgspec.Struct, rename="camel"):
    local: bool
    max_upload_size: int


class FixFontForm(msgspec.Struct):
    file: UploadFile
    theme: Annotated[str, msgspec.Meta(description="The theme definition (ThemeData) encoded as JSON.")]


class FixFontResult(msgspec.Struct, rename="camel"):
    filename: Annotated[str, msgspec.Meta(description="The suggested file name of the fixed pptx file.")]
    log: str
    download_url: Annotated[
        str,
        msgspec.Meta(
            description=(
                "The one-time URL of the fixed pptx file, which expires if not downloaded in time."
                " The filename query parameter overrides the suggested file name."
            )
        ),
    ]


class FontThemeRequest(msgspec.Struct):
    name: str
    theme: ThemeData


class InstallFontThemeRequest(msgspec.Struct):
    name: str
    theme: ThemeData
    overwrite: bool = False


class InstallFontThemeResult(msgspec.Struct):
    path: str


def _theme_error_extra(e: ThemeError) -> list[dict[str, Any]]:
    return [{"key": err.path, "message": err.message, "source": "body"} for err in e.errors]


def to_core_theme(data: ThemeData) -> Theme:
    """Validate the theme data with the same rules as the CLI, raising a 400 error for invalid values."""
    try:
        return load_theme(msgspec.to_builtins(data))
    except ThemeError as e:
        raise ValidationException(detail=str(e), extra=_theme_error_extra(e)) from e


def from_core_theme(theme_info: Theme) -> ThemeData:
    return msgspec.convert(dump_theme(theme_info), ThemeData)


def decode_theme_json(raw: str) -> Theme:
    """Decode and validate a JSON-encoded theme definition received as a form field."""
    try:
        data = msgspec.json.decode(raw, type=ThemeData)
    except msgspec.ValidationError as e:
        raise ValidationException(detail=f"Invalid theme: {e}") from e
    except msgspec.DecodeError as e:
        raise ValidationException(detail=f"The theme is not valid JSON: {e}") from e
    return to_core_theme(data)
