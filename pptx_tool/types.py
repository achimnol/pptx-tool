import dataclasses


@dataclasses.dataclass
class Theme:
    major_font_latin: str
    major_font_hangul: str
    major_font_symbol: str
    minor_font_latin: str
    minor_font_hangul: str
    minor_font_symbol: str
    title_bold: bool
    mono_font_latin: str = "Consolas"
    mono_font_hangul: str = "나눔고딕"
    body_first_level_style: str | None = None
    preserve_mono: bool = False
