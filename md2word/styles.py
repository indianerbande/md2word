"""Defines how the generated document looks (paragraph and character styles)."""

from __future__ import annotations

from typing import Any

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from md2word import docxutil as dx
from md2word.config import Config

# Style IDs the renderer uses.
# Deliberately identical to the names in Pandoc's reference document: that
# lets a Pandoc template be used via --reference-doc, and makes the return
# trip (docx -> Markdown) recognise code, quotes and captions again.
S_CODE_BLOCK = "SourceCode"
S_CODE_INLINE = "VerbatimChar"
S_QUOTE = "Quote"
S_CAPTION = "Caption"
S_FIGURE = "Figure"
S_HRULE = "HorizontalRule"
S_DEF_TERM = "DefinitionTerm"
S_DEF_BODY = "Definition"
S_TABLE_TEXT = "Compact"
S_META = "Author"
S_HYPERLINK = "Hyperlink"
S_FOOTNOTE_TEXT = "FootnoteText"
S_FOOTNOTE_REF = "FootnoteReference"

# Heading sizes as a factor of the base font size
_HEADING_SCALE = (1.85, 1.45, 1.25, 1.12, 1.03, 1.0)
_HEADING_SPACE_BEFORE = (18, 16, 13, 11, 9, 9)
_HEADING_SPACE_AFTER = (8, 7, 6, 5, 4, 4)


def hex_to_rgb(value: str) -> RGBColor:
    return RGBColor.from_string((value or "000000").lstrip("#").upper()[:6].rjust(6, "0"))


class StyleLookup:
    """Access to styles by their styleId.

    python-docx still supports lookup by ID but warns about it. This cache
    keeps the style objects to hand instead.
    """

    def __init__(self, document: Any) -> None:
        self._document = document
        self._cache: dict[str, Any] = {}
        self.refresh()

    def refresh(self) -> None:
        self._cache = {style.style_id: style for style in self._document.styles}

    def __getitem__(self, style_id: str) -> Any:
        style = self._cache.get(style_id)
        if style is None:
            self.refresh()
            style = self._cache.get(style_id)
        if style is None:
            raise KeyError(style_id)
        return style

    def get(self, style_id: str, default: Any = None) -> Any:
        try:
            return self[style_id]
        except KeyError:
            return default

    def heading(self, level: int) -> Any:
        return self[f"Heading{max(1, min(level, 9))}"]


# ----------------------------------------------------------------------
def _style_element(document: Any, style_id: str) -> Any:
    for node in document.styles.element.findall(qn("w:style")):
        if node.get(qn("w:styleId")) == style_id:
            return node
    return None


def ensure_style(
    document: Any,
    style_id: str,
    ui_name: str,
    style_type: WD_STYLE_TYPE,
    builtin: bool = False,
) -> Any:
    """Returns the style for a given styleId, creating it if necessary."""
    node = _style_element(document, style_id)
    if node is None:
        style = document.styles.add_style(ui_name, style_type, builtin=False)
        style.element.set(qn("w:styleId"), style_id)
        if builtin:
            style.element.attrib.pop(qn("w:customStyle"), None)
        node = style.element
    else:
        for style in document.styles:
            if style.style_id == style_id:
                return style
    for style in document.styles:
        if style.style_id == style_id:
            return style
    raise KeyError(style_id)


def _set_indent(style: Any, left: float = 0.0, right: float = 0.0, first_line: float = 0.0) -> None:
    fmt = style.paragraph_format
    fmt.left_indent = Mm(left)
    fmt.right_indent = Mm(right)
    if first_line:
        fmt.first_line_indent = Mm(first_line)


# ----------------------------------------------------------------------
class StyleBuilder:
    """Creates styles - and protects those of a reference template.

    When ``--reference-doc`` is in play, the styles defined there should
    win. :meth:`style` then returns ``None`` for styles that already exist,
    and the caller leaves them untouched.
    """

    def __init__(self, document: Any, config: Config, respect_existing: bool) -> None:
        self.doc = document
        self.cfg = config
        self.respect = respect_existing
        self.existing = {style.style_id for style in document.styles}

    def style(
        self,
        style_id: str,
        ui_name: str,
        style_type: WD_STYLE_TYPE,
        builtin: bool = False,
    ) -> Any:
        if style_id in self.existing and self.respect:
            return None
        return ensure_style(self.doc, style_id, ui_name, style_type, builtin)

    def existing_style(self, style_id: str) -> Any:
        """An existing style to edit - off limits with a reference template."""
        if self.respect:
            return None
        for style in self.doc.styles:
            if style.style_id == style_id:
                return style
        return None


def apply_styles(document: Any, config: Config, respect_existing: bool = False) -> None:
    """Sets up the page layout and every style according to the config."""
    builder = StyleBuilder(document, config, respect_existing)

    _setup_page(builder)
    _setup_normal(builder)
    _setup_headings(builder)
    _setup_title(builder)
    _setup_code(builder)
    _setup_quote(builder)
    _setup_misc(builder)
    _setup_footnotes(builder)

    if not respect_existing:
        dx.set_document_language(document, config.lang)
    dx.set_compat_mode(document)


def _setup_page(b: StyleBuilder) -> None:
    config = b.cfg
    # With a reference template its page layout stands, unless the call
    # explicitly asks for something else.
    explicit = config._explicit
    change_size = not b.respect or {"page_size", "landscape"} & explicit
    margin_fields = {"margin_top", "margin_bottom", "margin_left", "margin_right"}
    change_margins = not b.respect or margin_fields & explicit

    width, height = config.page_dimensions
    for section in b.doc.sections:
        if change_size:
            section.page_width = Mm(width)
            section.page_height = Mm(height)
        if change_margins:
            section.top_margin = Mm(config.margin_top)
            section.bottom_margin = Mm(config.margin_bottom)
            section.left_margin = Mm(config.margin_left)
            section.right_margin = Mm(config.margin_right)
            section.header_distance = Mm(max(8.0, config.margin_top / 2))
            section.footer_distance = Mm(max(8.0, config.margin_bottom / 2))


def _setup_normal(b: StyleBuilder) -> None:
    normal = b.existing_style("Normal")
    if normal is None:
        return
    config = b.cfg

    normal.font.name = config.body_font
    normal.font.size = Pt(config.font_size)
    normal.font.color.rgb = hex_to_rgb("000000")

    rPr = normal.element.get_or_add_rPr()
    fonts = dx.get_or_add(rPr, "w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(attr), config.body_font)

    fmt = normal.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(config.space_after)
    fmt.line_spacing = config.line_spacing
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    fmt.widow_control = True


def _setup_headings(b: StyleBuilder) -> None:
    config = b.cfg
    for level in range(1, 10):
        style = b.existing_style(f"Heading{level}")
        if style is None:
            continue

        index = min(level, len(_HEADING_SCALE)) - 1
        style.font.name = config.heading_font
        style.font.size = Pt(round(config.font_size * _HEADING_SCALE[index], 1))
        style.font.bold = level <= 3
        style.font.italic = level >= 5
        style.font.color.rgb = hex_to_rgb(config.heading_color)

        rPr = style.element.get_or_add_rPr()
        fonts = dx.get_or_add(rPr, "w:rFonts")
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            fonts.set(qn(attr), config.heading_font)

        fmt = style.paragraph_format
        fmt.space_before = Pt(_HEADING_SPACE_BEFORE[index])
        fmt.space_after = Pt(_HEADING_SPACE_AFTER[index])
        fmt.keep_with_next = True
        fmt.keep_together = True
        fmt.line_spacing = 1.0
        fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

        # A rule under H1/H2 only - it structures long documents
        if level <= 2:
            dx.set_borders(
                style.element.get_or_add_pPr(),
                "w:pBdr",
                edges=("bottom",),
                size=6 if level == 1 else 4,
                space=4,
                color=config.accent,
            )

    toc_heading = b.existing_style("TOCHeading")
    if toc_heading is None:
        return
    toc_heading.font.name = config.heading_font
    toc_heading.font.size = Pt(round(config.font_size * _HEADING_SCALE[0], 1))
    toc_heading.font.bold = True
    toc_heading.font.color.rgb = hex_to_rgb(config.heading_color)
    toc_heading.paragraph_format.space_after = Pt(12)


def _setup_title(b: StyleBuilder) -> None:
    config = b.cfg

    title = b.existing_style("Title")
    if title is not None:
        title.font.name = config.heading_font
        title.font.size = Pt(round(config.font_size * 2.6, 1))
        title.font.bold = True
        title.font.color.rgb = hex_to_rgb(config.heading_color)
        title.paragraph_format.space_after = Pt(6)
        title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = b.existing_style("Subtitle")
    if subtitle is not None:
        subtitle.font.name = config.heading_font
        subtitle.font.size = Pt(round(config.font_size * 1.4, 1))
        subtitle.font.italic = True
        subtitle.font.color.rgb = hex_to_rgb(config.quote_color)
        subtitle.paragraph_format.space_after = Pt(24)
        subtitle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _setup_code(b: StyleBuilder) -> None:
    config = b.cfg

    block = b.style(S_CODE_BLOCK, "Source Code", WD_STYLE_TYPE.PARAGRAPH)
    if block is not None:
        block.font.name = config.code_font
        block.font.size = Pt(config.code_font_size)
        rPr = block.element.get_or_add_rPr()
        fonts = dx.get_or_add(rPr, "w:rFonts")
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            fonts.set(qn(attr), config.code_font)
        dx.get_or_add(rPr, "w:noProof").set(qn("w:val"), "1")

        fmt = block.paragraph_format
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = 1.0
        fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        fmt.keep_together = True
        _set_indent(block, left=3.0, right=2.0)
        dx.set_shading(block.element.get_or_add_pPr(), config.code_bg)
        # The renderer sets the borders per code block (outer edges only)

    inline = b.style(S_CODE_INLINE, "Verbatim Char", WD_STYLE_TYPE.CHARACTER)
    if inline is not None:
        inline.font.name = config.code_font
        inline.font.size = Pt(config.code_font_size)
        rPr = inline.element.get_or_add_rPr()
        fonts = dx.get_or_add(rPr, "w:rFonts")
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            fonts.set(qn(attr), config.code_font)
        dx.set_shading(rPr, config.code_bg)
        dx.get_or_add(rPr, "w:noProof").set(qn("w:val"), "1")


def _setup_quote(b: StyleBuilder) -> None:
    config = b.cfg
    quote = b.style(S_QUOTE, "Quote", WD_STYLE_TYPE.PARAGRAPH, builtin=True)
    if quote is None:
        return
    quote.font.color.rgb = hex_to_rgb(config.quote_color)
    quote.font.italic = True
    fmt = quote.paragraph_format
    fmt.space_before = Pt(6)
    fmt.space_after = Pt(6)
    _set_indent(quote, left=8.0, right=4.0)
    dx.set_borders(
        quote.element.get_or_add_pPr(),
        "w:pBdr",
        edges=("left",),
        size=18,
        space=10,
        color=config.accent,
    )


def _setup_misc(b: StyleBuilder) -> None:
    config = b.cfg

    caption = b.style(S_CAPTION, "Caption", WD_STYLE_TYPE.PARAGRAPH, builtin=True)
    if caption is not None:
        caption.font.size = Pt(max(7.0, config.font_size - 2))
        caption.font.italic = True
        caption.font.color.rgb = hex_to_rgb(config.quote_color)
        caption.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.paragraph_format.space_before = Pt(2)
        caption.paragraph_format.space_after = Pt(10)

    figure = b.style(S_FIGURE, "Figure", WD_STYLE_TYPE.PARAGRAPH)
    if figure is not None:
        figure.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        figure.paragraph_format.space_before = Pt(8)
        figure.paragraph_format.space_after = Pt(4)
        figure.paragraph_format.keep_with_next = True

    rule = b.style(S_HRULE, "Horizontal Rule", WD_STYLE_TYPE.PARAGRAPH)
    if rule is not None:
        rule.font.size = Pt(1)
        rule.paragraph_format.space_before = Pt(8)
        rule.paragraph_format.space_after = Pt(10)
        dx.set_borders(
            rule.element.get_or_add_pPr(),
            "w:pBdr",
            edges=("bottom",),
            size=6,
            space=1,
            color=config.accent,
        )

    term = b.style(S_DEF_TERM, "Definition Term", WD_STYLE_TYPE.PARAGRAPH)
    if term is not None:
        term.font.bold = True
        term.paragraph_format.space_before = Pt(6)
        term.paragraph_format.space_after = Pt(2)
        term.paragraph_format.keep_with_next = True

    body = b.style(S_DEF_BODY, "Definition", WD_STYLE_TYPE.PARAGRAPH)
    if body is not None:
        body.paragraph_format.space_after = Pt(4)
        _set_indent(body, left=8.0)

    # Text inside table cells: more compact than body text
    cell = b.style(S_TABLE_TEXT, "Compact", WD_STYLE_TYPE.PARAGRAPH)
    if cell is not None:
        cell.font.size = Pt(max(8.0, config.font_size - 1))
        cell.paragraph_format.space_before = Pt(1)
        cell.paragraph_format.space_after = Pt(1)
        cell.paragraph_format.line_spacing = 1.0
        cell.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

    # Author and date line on the title page
    meta = b.style(S_META, "Author", WD_STYLE_TYPE.PARAGRAPH)
    if meta is not None:
        meta.font.size = Pt(config.font_size)
        meta.font.color.rgb = hex_to_rgb(config.quote_color)
        meta.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        meta.paragraph_format.space_after = Pt(2)

    link = b.style(S_HYPERLINK, "Hyperlink", WD_STYLE_TYPE.CHARACTER, builtin=True)
    if link is not None:
        link.font.color.rgb = hex_to_rgb(config.link_color)
        link.font.underline = True


def _setup_footnotes(b: StyleBuilder) -> None:
    config = b.cfg

    text = b.style(S_FOOTNOTE_TEXT, "footnote text", WD_STYLE_TYPE.PARAGRAPH, builtin=True)
    if text is not None:
        text.font.size = Pt(max(7.5, config.font_size - 2))
        fmt = text.paragraph_format
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(2)
        fmt.line_spacing = 1.0
        fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

    ref = b.style(S_FOOTNOTE_REF, "footnote reference", WD_STYLE_TYPE.CHARACTER, builtin=True)
    if ref is not None:
        dx.get_or_add(ref.element.get_or_add_rPr(), "w:vertAlign").set(
            qn("w:val"), "superscript"
        )


# ----------------------------------------------------------------------
def build_table_style(document: Any, config: Config) -> str:
    """Returns the styleId of the table style to use."""
    available = {style.style_id for style in document.styles}
    return "TableGrid" if "TableGrid" in available else "TableNormal"


def apply_table_alignment(table: Any) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
