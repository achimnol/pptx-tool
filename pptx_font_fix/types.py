import dataclasses
from typing import Optional


@dataclasses.dataclass
class Theme:
    major_font_latin: str
    major_font_hangul: str
    major_font_symbol: str
    minor_font_latin: str
    minor_font_hangul: str
    minor_font_symbol: str
    title_bold: bool
    body_first_level_style: Optional[str]