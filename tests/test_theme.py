import copy
import json
from pathlib import Path
from typing import Any

import pytest

from pptx_tool.theme import ThemeError, dump_theme, list_bundled_themes, load_theme, load_theme_file, resolve_theme_arg

VALID_THEME: dict[str, Any] = {
    "majorFont": {"latin": "A", "hangul": "B", "symbol": "C"},
    "minorFont": {"latin": "D", "hangul": "E", "symbol": "F"},
    "monoFont": {"latin": "G", "hangul": "H"},
    "options": {"titleBold": True, "bodyFirstLevelStyle": None},
}


def _with(path: str, value: Any) -> dict[str, Any]:
    """Return a copy of VALID_THEME with the dotted path set to value, or removed if value is ..."""
    data = copy.deepcopy(VALID_THEME)
    *parents, key = path.split(".")
    target = data
    for parent in parents:
        target = target[parent]
    if value is ...:
        del target[key]
    else:
        target[key] = value
    return data


def test_load_theme_defaults_preserve_mono() -> None:
    theme_info = load_theme(VALID_THEME)
    assert theme_info.major_font_latin == "A"
    assert theme_info.mono_font_hangul == "H"
    assert theme_info.body_first_level_style is None
    assert theme_info.preserve_mono is False


def test_dump_theme_round_trip() -> None:
    data = _with("options.bodyFirstLevelStyle", "Custom Weight")
    data["options"]["preserveMono"] = True
    dumped = dump_theme(load_theme(data))
    assert dumped == data
    assert load_theme(dumped) == load_theme(data)


def test_dump_theme_always_includes_preserve_mono() -> None:
    assert dump_theme(load_theme(VALID_THEME))["options"]["preserveMono"] is False


@pytest.mark.parametrize(
    "path,value",
    [
        ("majorFont.latin", ...),
        ("majorFont.latin", ""),
        ("majorFont.latin", "   "),
        ("minorFont.symbol", 3),
        ("monoFont.hangul", None),
        ("options.titleBold", "true"),
        ("options.titleBold", ...),
        ("options.bodyFirstLevelStyle", ...),
        ("options.bodyFirstLevelStyle", ""),
        ("options.preserveMono", 1),
    ],
)
def test_load_theme_reports_field_path(path: str, value: Any) -> None:
    with pytest.raises(ThemeError) as exc_info:
        load_theme(_with(path, value))
    assert [e.path for e in exc_info.value.errors] == [path]


def test_load_theme_reports_missing_objects() -> None:
    data = _with("minorFont", ...)
    data["options"] = []
    with pytest.raises(ThemeError) as exc_info:
        load_theme(data)
    paths = [e.path for e in exc_info.value.errors]
    assert paths[:4] == ["minorFont", "minorFont.latin", "minorFont.hangul", "minorFont.symbol"]
    assert "options" in paths
    assert "options.titleBold" in paths


def test_load_theme_rejects_non_object() -> None:
    with pytest.raises(ThemeError) as exc_info:
        load_theme([])  # type: ignore[arg-type]
    assert [e.path for e in exc_info.value.errors] == [""]


def test_load_theme_file_rejects_invalid_json(tmp_path: Path) -> None:
    path = tmp_path / "broken.json"
    path.write_text("{")
    with pytest.raises(ThemeError, match="not a valid JSON file"):
        load_theme_file(path)


def test_load_theme_file(tmp_path: Path) -> None:
    path = tmp_path / "theme.json"
    path.write_text(json.dumps(VALID_THEME))
    assert load_theme_file(path) == load_theme(VALID_THEME)


def test_list_bundled_themes() -> None:
    themes = list_bundled_themes()
    assert len(themes) >= 11
    ids = [t.id for t in themes]
    assert ids == sorted(ids)
    pretendard = next(t for t in themes if t.id == "pretendard")
    assert pretendard.name == "Pretendard"
    assert pretendard.theme.major_font_latin == "Pretendard"


def test_resolve_theme_arg_bundled_id() -> None:
    assert resolve_theme_arg("pretendard").major_font_latin == "Pretendard"


def test_resolve_theme_arg_path_wins(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    Path("pretendard").write_text(json.dumps(VALID_THEME))
    assert resolve_theme_arg("pretendard").major_font_latin == "A"


def test_resolve_theme_arg_unknown() -> None:
    with pytest.raises(ThemeError, match="neither a theme file nor a bundled theme"):
        resolve_theme_arg("no-such-theme")
