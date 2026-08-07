"""Configuration: page layout, typography and conversion options."""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Any

# Paper sizes in millimetres (width x height, portrait)
PAGE_SIZES: dict[str, tuple[float, float]] = {
    "a4": (210.0, 297.0),
    "a5": (148.0, 210.0),
    "a3": (297.0, 420.0),
    "letter": (215.9, 279.4),
    "legal": (215.9, 355.6),
}

# Predefined colour and font schemes (hex without the '#')
THEMES: dict[str, dict[str, Any]] = {
    "default": {
        "body_font": "Calibri",
        "heading_font": "Calibri Light",
        "code_font": "Consolas",
        "accent": "2F5496",
        "heading_color": "2F5496",
        "link_color": "0563C1",
        "code_bg": "F5F5F5",
        "quote_color": "5A5A5A",
        "table_header_bg": "D9E2F3",
    },
    "classic": {
        "body_font": "Times New Roman",
        "heading_font": "Times New Roman",
        "code_font": "Courier New",
        "accent": "000000",
        "heading_color": "000000",
        "link_color": "0000EE",
        "code_bg": "F2F2F2",
        "quote_color": "404040",
        "table_header_bg": "E6E6E6",
    },
    "modern": {
        "body_font": "Segoe UI",
        "heading_font": "Segoe UI Semibold",
        "code_font": "Cascadia Mono",
        "accent": "0F6CBD",
        "heading_color": "13343B",
        "link_color": "0F6CBD",
        "code_bg": "F3F6F9",
        "quote_color": "4A5A66",
        "table_header_bg": "DCE9F5",
    },
    "mono": {
        "body_font": "Arial",
        "heading_font": "Arial Black",
        "code_font": "Menlo",
        "accent": "222222",
        "heading_color": "111111",
        "link_color": "1A1A1A",
        "code_bg": "EFEFEF",
        "quote_color": "555555",
        "table_header_bg": "DDDDDD",
    },
}

DEFAULT_THEME = "default"


@dataclass
class Config:
    """Every knob of the conversion in one place."""

    # --- Page ---------------------------------------------------------
    page_size: str = "a4"
    landscape: bool = False
    margin_top: float = 25.0  # mm
    margin_bottom: float = 25.0
    margin_left: float = 25.0
    margin_right: float = 25.0

    # --- Typography ---------------------------------------------------
    theme: str = DEFAULT_THEME
    body_font: str = ""
    heading_font: str = ""
    code_font: str = ""
    font_size: float = 11.0  # pt
    code_font_size: float = 9.5
    line_spacing: float = 1.15
    space_after: float = 8.0  # pt Abstand nach Absaetzen

    # --- Colours (hex without the '#') ---------------------------------
    accent: str = ""
    heading_color: str = ""
    link_color: str = ""
    code_bg: str = ""
    quote_color: str = ""
    table_header_bg: str = ""

    # --- Structure ------------------------------------------------------
    toc: bool = False
    toc_depth: int = 3
    # Empty = derive from the document language (see i18n.translate)
    toc_title: str = ""
    title_page: bool = False
    number_headings: bool = False
    page_numbers: bool = False
    header_text: str = ""
    footer_text: str = ""
    break_on_h1: bool = False

    # --- Content --------------------------------------------------------
    highlight: bool = True
    pygments_style: str = "friendly"
    max_image_width: float = 0.0  # mm, 0 = Textbreite
    download_images: bool = True
    image_timeout: float = 10.0
    captions: str = "title"  # title | alt | none
    strip_html: bool = False
    math_mode: str = "omml"  # omml | text
    footnote_mode: str = "footnotes"  # footnotes | endnotes
    lang: str = "en-US"

    # --- Document metadata (may come from the front matter) -------------
    title: str = ""
    subtitle: str = ""
    author: str = ""
    date: str = ""
    subject: str = ""
    keywords: str = ""
    comments: str = ""

    # --- Miscellaneous --------------------------------------------------
    reference_doc: str = ""
    base_dir: str = "."
    verbose: bool = False

    # Front-matter keys that are not metadata
    _extra: dict[str, Any] = field(default_factory=dict, repr=False)
    # Fields set explicitly on the command line, which the front matter
    # is therefore not allowed to override
    _explicit: set = field(default_factory=set, repr=False)

    def __post_init__(self) -> None:
        self.apply_theme(self.theme)

    # ------------------------------------------------------------------
    def apply_theme(self, name: str) -> None:
        """Fills every colour and font field still empty from the theme."""
        palette = THEMES.get(name, THEMES[DEFAULT_THEME])
        for key, value in palette.items():
            if not getattr(self, key, ""):
                setattr(self, key, value)

    @property
    def page_dimensions(self) -> tuple[float, float]:
        """Page width and height in mm, honouring the orientation."""
        width, height = PAGE_SIZES.get(self.page_size, PAGE_SIZES["a4"])
        if self.landscape:
            width, height = height, width
        return width, height

    @property
    def text_width_mm(self) -> float:
        """Usable text width in mm."""
        width, _ = self.page_dimensions
        return max(10.0, width - self.margin_left - self.margin_right)

    def merged_with(self, **overrides: Any) -> "Config":
        """A copy with overridden fields (empty values are ignored)."""
        clean = {k: v for k, v in overrides.items() if v not in ("", None)}
        return replace(self, **clean)
