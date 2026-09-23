"""Parse, validate and serialize the theme JSON files shared by the CLI and the web UI."""

import dataclasses
import importlib.resources
import json
from collections.abc import Mapping
from importlib.resources.abc import Traversable
from pathlib import Path
from typing import Any

from .types import Theme

_FONT_SETS: dict[str, tuple[str, ...]] = {
    "majorFont": ("latin", "hangul", "symbol"),
    "minorFont": ("latin", "hangul", "symbol"),
    "monoFont": ("latin", "hangul"),
}


@dataclasses.dataclass(frozen=True)
class ThemeFieldError:
    path: str
    """The dotted JSON path of the invalid field, such as "majorFont.latin", or empty for the whole document."""
    message: str

    def __str__(self) -> str:
        return f"{self.path}: {self.message}" if self.path else self.message


class ThemeError(ValueError):
    """Raised when a theme definition is invalid, carrying all field errors found."""

    def __init__(self, errors: list[ThemeFieldError]) -> None:
        super().__init__(errors)
        self.errors = errors

    def __str__(self) -> str:
        return "invalid theme: " + "; ".join(str(e) for e in self.errors)


def load_theme(data: Mapping[str, Any]) -> Theme:
    """
    Build a Theme from the camelCase theme JSON structure.

    All keys are required except `options.preserveMono`, which defaults to false.
    `options.bodyFirstLevelStyle` must be present but may be null.
    Unknown keys are ignored, including the top-level `name` of the bundled themes.
    """
    errors: list[ThemeFieldError] = []

    def _get_object(parent: Mapping[str, Any], key: str, path: str) -> Mapping[str, Any]:
        value = parent.get(key)
        if not isinstance(value, Mapping):
            errors.append(ThemeFieldError(path, "must be an object"))
            return {}
        return value

    def _get_font(parent: Mapping[str, Any], key: str, path: str) -> str:
        value = parent.get(key)
        if not isinstance(value, str) or not value.strip():
            errors.append(ThemeFieldError(path, "must be a non-empty string"))
            return ""
        return value

    if not isinstance(data, Mapping):
        raise ThemeError([ThemeFieldError("", "the theme must be a JSON object")])

    fonts: dict[str, str] = {}
    for set_name, scripts in _FONT_SETS.items():
        font_set = _get_object(data, set_name, set_name)
        for script in scripts:
            fonts[f"{set_name}.{script}"] = _get_font(font_set, script, f"{set_name}.{script}")

    options = _get_object(data, "options", "options")
    title_bold = options.get("titleBold")
    if not isinstance(title_bold, bool):
        errors.append(ThemeFieldError("options.titleBold", "must be a boolean"))
    if "bodyFirstLevelStyle" not in options:
        errors.append(ThemeFieldError("options.bodyFirstLevelStyle", "must be present (use null to disable it)"))
    body_first_level_style = options.get("bodyFirstLevelStyle")
    if body_first_level_style is not None and not (
        isinstance(body_first_level_style, str) and body_first_level_style.strip()
    ):
        errors.append(ThemeFieldError("options.bodyFirstLevelStyle", "must be a non-empty string or null"))
    preserve_mono = options.get("preserveMono", False)
    if not isinstance(preserve_mono, bool):
        errors.append(ThemeFieldError("options.preserveMono", "must be a boolean"))

    if errors:
        raise ThemeError(errors)
    return Theme(
        major_font_latin=fonts["majorFont.latin"],
        major_font_hangul=fonts["majorFont.hangul"],
        major_font_symbol=fonts["majorFont.symbol"],
        minor_font_latin=fonts["minorFont.latin"],
        minor_font_hangul=fonts["minorFont.hangul"],
        minor_font_symbol=fonts["minorFont.symbol"],
        mono_font_latin=fonts["monoFont.latin"],
        mono_font_hangul=fonts["monoFont.hangul"],
        title_bold=bool(title_bold),
        body_first_level_style=body_first_level_style,
        preserve_mono=bool(preserve_mono),
    )


def dump_theme(theme_info: Theme) -> dict[str, Any]:
    """Serialize a Theme into the camelCase theme JSON structure accepted by `load_theme()`.

    The display name of a bundled theme is not part of a Theme, so it is never included.
    """
    return {
        "majorFont": {
            "latin": theme_info.major_font_latin,
            "hangul": theme_info.major_font_hangul,
            "symbol": theme_info.major_font_symbol,
        },
        "minorFont": {
            "latin": theme_info.minor_font_latin,
            "hangul": theme_info.minor_font_hangul,
            "symbol": theme_info.minor_font_symbol,
        },
        "monoFont": {
            "latin": theme_info.mono_font_latin,
            "hangul": theme_info.mono_font_hangul,
        },
        "options": {
            "titleBold": theme_info.title_bold,
            "bodyFirstLevelStyle": theme_info.body_first_level_style,
            "preserveMono": theme_info.preserve_mono,
        },
    }


def load_theme_file(path: Path) -> Theme:
    """Read and validate a theme JSON file."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise ThemeError([ThemeFieldError("", f"{path} is not a valid JSON file ({e})")]) from e
    return load_theme(data)


@dataclasses.dataclass(frozen=True)
class BundledTheme:
    id: str
    name: str
    theme: Theme


def _bundled_theme_dir() -> Traversable:
    return importlib.resources.files("pptx_tool") / "themes"


def list_bundled_themes() -> list[BundledTheme]:
    """List the themes shipped with the package, sorted by their ids (file stems).

    Each bundled theme file carries its display name in the top-level `name` key.
    """
    themes: list[BundledTheme] = []
    for entry in _bundled_theme_dir().iterdir():
        if not entry.name.endswith(".json"):
            continue
        theme_id = entry.name.removesuffix(".json")
        data = json.loads(entry.read_text(encoding="utf-8"))
        name = data.get("name") if isinstance(data, Mapping) else None
        if not isinstance(name, str) or not name.strip():
            raise ThemeError([
                ThemeFieldError("name", f"bundled theme {theme_id!r} must define a non-empty display name"),
            ])
        themes.append(BundledTheme(theme_id, name, load_theme(data)))
    return sorted(themes, key=lambda t: t.id)


def resolve_theme_arg(value: str) -> Theme:
    """Load a theme from a JSON file path, or from a bundled theme id such as "pretendard"."""
    path = Path(value)
    if path.is_file():
        return load_theme_file(path)
    bundled = {t.id: t for t in list_bundled_themes()}
    if value in bundled:
        return bundled[value].theme
    raise ThemeError([
        ThemeFieldError("", f"{value!r} is neither a theme file nor a bundled theme id ({', '.join(sorted(bundled))})")
    ])
