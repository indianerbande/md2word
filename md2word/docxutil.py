"""Low-Level-Helfer fuer OOXML-Konstrukte, die python-docx nicht direkt anbietet.

Hier liegt alles, was direkt am XML arbeitet: Schattierungen, Rahmen,
Hyperlinks, Lesezeichen, Feldfunktionen (Inhaltsverzeichnis, Seitenzahlen),
mehrstufige Listen-Nummerierungen und echte Word-Fussnoten.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Any, Iterable

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.ns import nsmap, qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.shared import Pt

W = nsmap["w"]
R = nsmap["r"]

XML_DECL = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
)
NS_DECL = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in nsmap.items())


# ----------------------------------------------------------------------
# Allgemeine Element-Helfer
# ----------------------------------------------------------------------
def make_element(tag: str, **attrs: Any) -> Any:
    """Erzeugt ein OOXML-Element mit w:-Attributen."""
    element = OxmlElement(tag)
    for key, value in attrs.items():
        element.set(qn(key.replace("__", ":")), str(value))
    return element


def get_or_add(parent: Any, tag: str, position: int | None = None) -> Any:
    """Liefert das erste Kindelement `tag` oder legt es an."""
    found = parent.find(qn(tag))
    if found is None:
        found = OxmlElement(tag)
        if position is None:
            parent.append(found)
        else:
            parent.insert(position, found)
    return found


def set_shading(element: Any, color: str, fill: str | None = None) -> None:
    """Setzt eine Hintergrundfarbe (w:shd) auf pPr/rPr/tcPr."""
    shd = get_or_add(element, "w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill or color)


def set_borders(
    element: Any,
    tag: str = "w:pBdr",
    edges: Iterable[str] = ("top", "left", "bottom", "right"),
    style: str = "single",
    size: int = 6,
    space: int = 4,
    color: str = "auto",
) -> None:
    """Setzt Rahmenlinien auf einen Absatz (pBdr) oder eine Zelle (tcBorders)."""
    borders = get_or_add(element, tag)
    order = ["top", "left", "bottom", "right", "insideH", "insideV"]
    wanted = set(edges)
    for edge in order:
        existing = borders.find(qn(f"w:{edge}"))
        if edge not in wanted:
            continue
        node = existing if existing is not None else OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), style)
        node.set(qn("w:sz"), str(size))
        node.set(qn("w:space"), str(space))
        node.set(qn("w:color"), color)
        if existing is None:
            borders.append(node)


def paragraph_properties(paragraph: Any) -> Any:
    """Liefert das pPr-Element eines Absatzes (legt es bei Bedarf an)."""
    return paragraph._p.get_or_add_pPr()


def shade_paragraph(paragraph: Any, fill: str) -> None:
    set_shading(paragraph_properties(paragraph), fill)


def set_paragraph_border(paragraph: Any, **kwargs: Any) -> None:
    set_borders(paragraph_properties(paragraph), "w:pBdr", **kwargs)


def keep_with_next(paragraph: Any, value: bool = True) -> None:
    """Verhindert einen Seitenumbruch zwischen diesem und dem naechsten Absatz."""
    pPr = paragraph_properties(paragraph)
    node = get_or_add(pPr, "w:keepNext")
    node.set(qn("w:val"), "1" if value else "0")


def keep_lines_together(paragraph: Any) -> None:
    node = get_or_add(paragraph_properties(paragraph), "w:keepLines")
    node.set(qn("w:val"), "1")


def prevent_widows(paragraph: Any) -> None:
    get_or_add(paragraph_properties(paragraph), "w:widowControl").set(qn("w:val"), "1")


def set_cell_background(cell: Any, fill: str) -> None:
    set_shading(cell._tc.get_or_add_tcPr(), fill)


def set_cell_vertical_alignment(cell: Any, value: str = "center") -> None:
    node = get_or_add(cell._tc.get_or_add_tcPr(), "w:vAlign")
    node.set(qn("w:val"), value)


def add_page_break(paragraph: Any) -> None:
    """Haengt einen harten Seitenumbruch an einen Absatz."""
    run = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run.append(br)
    paragraph._p.append(run)


def set_run_font(run: Any, name: str) -> None:
    """Setzt die Schriftart inkl. East-Asian-/Complex-Script-Varianten."""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rFonts = get_or_add(rPr, "w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)


def set_run_style(run: Any, style_id: str) -> None:
    rPr = run._r.get_or_add_rPr()
    node = get_or_add(rPr, "w:rStyle")
    node.set(qn("w:val"), style_id)


def set_vertical_align(run: Any, value: str) -> None:
    """value: 'superscript' | 'subscript' | 'baseline'."""
    rPr = run._r.get_or_add_rPr()
    node = get_or_add(rPr, "w:vertAlign")
    node.set(qn("w:val"), value)


def set_no_proof(run: Any) -> None:
    """Schaltet Rechtschreib-/Grammatikpruefung fuer den Run ab (z. B. Code)."""
    get_or_add(run._r.get_or_add_rPr(), "w:noProof").set(qn("w:val"), "1")


# ----------------------------------------------------------------------
# Lesezeichen und Hyperlinks
# ----------------------------------------------------------------------
_slug_counter: dict[str, int] = {}


def slugify(text: str, existing: set[str] | None = None) -> str:
    """Erzeugt eine GitHub-aehnliche Anker-ID aus einem Ueberschriftentext."""
    value = unicodedata.normalize("NFKD", text or "")
    value = value.replace("ß", "ss")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^\w\s-]", "", value, flags=re.U).strip().lower()
    value = re.sub(r"[\s_]+", "-", value)
    value = value.strip("-") or "abschnitt"

    if existing is not None:
        base, index = value, 1
        while value in existing:
            value = f"{base}-{index}"
            index += 1
        existing.add(value)
    return value


def sanitize_bookmark(name: str) -> str:
    """Word-Lesezeichen: max. 40 Zeichen, Buchstaben/Ziffern/Unterstrich."""
    clean = re.sub(r"[^\w]", "_", name, flags=re.U)
    if not clean or clean[0].isdigit():
        clean = "_" + clean
    return clean[:40]


def add_bookmark(paragraph: Any, name: str, bookmark_id: int) -> None:
    """Umschliesst den Absatzinhalt mit einem Lesezeichen."""
    safe = sanitize_bookmark(name)
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), safe)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))

    pPr = paragraph._p.find(qn("w:pPr"))
    paragraph._p.insert(1 if pPr is not None else 0, start)
    paragraph._p.append(end)


def add_external_hyperlink(paragraph: Any, url: str) -> Any:
    """Legt ein leeres w:hyperlink-Element mit externer Beziehung an."""
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    paragraph._p.append(link)
    return link


def add_internal_hyperlink(paragraph: Any, anchor: str) -> Any:
    """Legt ein w:hyperlink-Element auf ein Lesezeichen im Dokument an."""
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), sanitize_bookmark(anchor))
    paragraph._p.append(link)
    return link


def move_run_into(container: Any, run: Any) -> None:
    """Verschiebt einen bereits erzeugten Run in ein Hyperlink-Element."""
    run._r.getparent().remove(run._r)
    container.append(run._r)


# ----------------------------------------------------------------------
# Feldfunktionen (TOC, PAGE, NUMPAGES, ...)
# ----------------------------------------------------------------------
def add_field(paragraph: Any, instruction: str, placeholder: str = "", dirty: bool = True) -> None:
    """Fuegt eine Word-Feldfunktion als Run-Sequenz in einen Absatz ein."""
    begin = OxmlElement("w:r")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    if dirty:
        fld_begin.set(qn("w:dirty"), "true")
    begin.append(fld_begin)

    instr_run = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction.strip()} "
    instr_run.append(instr)

    separate = OxmlElement("w:r")
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    separate.append(fld_sep)

    result = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.set(qn("xml:space"), "preserve")
    text.text = placeholder
    result.append(text)

    end = OxmlElement("w:r")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    end.append(fld_end)

    for node in (begin, instr_run, separate, result, end):
        paragraph._p.append(node)


def force_field_update_on_open(document: Any) -> None:
    """Setzt <w:updateFields/>, damit Word Felder beim Oeffnen neu berechnet."""
    settings = document.settings.element
    node = settings.find(qn("w:updateFields"))
    if node is None:
        node = OxmlElement("w:updateFields")
        settings.insert(0, node)
    node.set(qn("w:val"), "true")


# ----------------------------------------------------------------------
# Mehrstufige Listen-Nummerierung
# ----------------------------------------------------------------------
_BULLET_GLYPHS = [("", "Symbol"), ("o", "Courier New"), ("", "Wingdings")]
_NUMBER_FORMATS = ["decimal", "lowerLetter", "lowerRoman"]


class NumberingRegistry:
    """Verwaltet abstractNum-/num-Definitionen in word/numbering.xml."""

    def __init__(self, document: Any) -> None:
        self._numbering = document.part.numbering_part.element
        self._bullet_abstract: int | None = None
        self._decimal_abstract: int | None = None

    # -- interne Helfer -------------------------------------------------
    def _next_abstract_id(self) -> int:
        ids = [
            int(node.get(qn("w:abstractNumId")))
            for node in self._numbering.findall(qn("w:abstractNum"))
            if node.get(qn("w:abstractNumId")) is not None
        ]
        return max(ids, default=-1) + 1

    def _next_num_id(self) -> int:
        ids = [
            int(node.get(qn("w:numId")))
            for node in self._numbering.findall(qn("w:num"))
            if node.get(qn("w:numId")) is not None
        ]
        return max(ids, default=0) + 1

    def _insert_abstract(self, element: Any) -> None:
        """abstractNum-Elemente muessen laut Schema vor den num-Elementen stehen."""
        existing = self._numbering.findall(qn("w:abstractNum"))
        if existing:
            existing[-1].addnext(element)
            return
        first_num = self._numbering.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(element)
        else:
            self._numbering.append(element)

    def _build_abstract(self, ordered: bool) -> int:
        abstract_id = self._next_abstract_id()
        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(abstract_id))

        multi = OxmlElement("w:multiLevelType")
        multi.set(qn("w:val"), "hybridMultilevel")
        abstract.append(multi)

        for level in range(9):
            abstract.append(self._build_level(level, ordered))

        self._insert_abstract(abstract)
        return abstract_id

    @staticmethod
    def _build_level(level: int, ordered: bool) -> Any:
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(level))

        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        lvl.append(start)

        fmt = OxmlElement("w:numFmt")
        text = OxmlElement("w:lvlText")
        if ordered:
            fmt.set(qn("w:val"), _NUMBER_FORMATS[level % len(_NUMBER_FORMATS)])
            text.set(qn("w:val"), f"%{level + 1}.")
        else:
            glyph, _font = _BULLET_GLYPHS[level % len(_BULLET_GLYPHS)]
            fmt.set(qn("w:val"), "bullet")
            text.set(qn("w:val"), glyph)
        lvl.append(fmt)
        lvl.append(text)

        justify = OxmlElement("w:lvlJc")
        justify.set(qn("w:val"), "left")
        lvl.append(justify)

        pPr = OxmlElement("w:pPr")
        indent = OxmlElement("w:ind")
        indent.set(qn("w:left"), str(720 * (level + 1)))
        indent.set(qn("w:hanging"), "360")
        pPr.append(indent)
        lvl.append(pPr)

        if not ordered:
            _glyph, font = _BULLET_GLYPHS[level % len(_BULLET_GLYPHS)]
            rPr = OxmlElement("w:rPr")
            rFonts = OxmlElement("w:rFonts")
            rFonts.set(qn("w:ascii"), font)
            rFonts.set(qn("w:hAnsi"), font)
            rFonts.set(qn("w:hint"), "default")
            rPr.append(rFonts)
            lvl.append(rPr)

        return lvl

    def _create_num(self, abstract_id: int, start_at: int = 1) -> int:
        num_id = self._next_num_id()
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        ref = OxmlElement("w:abstractNumId")
        ref.set(qn("w:val"), str(abstract_id))
        num.append(ref)

        if start_at != 1:
            override = OxmlElement("w:lvlOverride")
            override.set(qn("w:ilvl"), "0")
            start_override = OxmlElement("w:startOverride")
            start_override.set(qn("w:val"), str(start_at))
            override.append(start_override)
            num.append(override)

        self._numbering.append(num)
        return num_id

    # -- oeffentliche API ------------------------------------------------
    def new_list(self, ordered: bool, start_at: int = 1) -> int:
        """Liefert eine numId fuer eine neue Liste (Nummerierung startet neu)."""
        if ordered:
            if self._decimal_abstract is None:
                self._decimal_abstract = self._build_abstract(ordered=True)
            return self._create_num(self._decimal_abstract, start_at)

        if self._bullet_abstract is None:
            self._bullet_abstract = self._build_abstract(ordered=False)
        return self._create_num(self._bullet_abstract)


MAX_LIST_LEVEL = 8  # Word kennt die Ebenen 0 bis 8


def apply_numbering(paragraph: Any, num_id: int, level: int) -> None:
    """Haengt einen Absatz an eine Listendefinition."""
    pPr = paragraph_properties(paragraph)
    numPr = get_or_add(pPr, "w:numPr")
    level = max(0, min(level, MAX_LIST_LEVEL))
    for tag, value in (("w:ilvl", level), ("w:numId", num_id)):
        node = get_or_add(numPr, tag)
        node.set(qn("w:val"), str(value))


# ----------------------------------------------------------------------
# Ueberschriften-Nummerierung (1., 1.1, 1.1.1 ...)
# ----------------------------------------------------------------------
def enable_heading_numbering(document: Any, depth: int = 3) -> None:
    """Verknuepft die Ueberschriftenformate mit einer mehrstufigen Nummerierung."""
    numbering = document.part.numbering_part.element
    registry = NumberingRegistry(document)
    abstract_id = registry._next_abstract_id()

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)

    for level in range(9):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(level))

        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        lvl.append(start)

        fmt = OxmlElement("w:numFmt")
        fmt.set(qn("w:val"), "decimal" if level < depth else "none")
        lvl.append(fmt)

        style_ref = OxmlElement("w:pStyle")
        style_ref.set(qn("w:val"), f"Heading{level + 1}")
        lvl.append(style_ref)

        text = OxmlElement("w:lvlText")
        text.set(qn("w:val"), ".".join(f"%{i + 1}" for i in range(level + 1)) if level < depth else "")
        lvl.append(text)

        justify = OxmlElement("w:lvlJc")
        justify.set(qn("w:val"), "left")
        lvl.append(justify)

        pPr = OxmlElement("w:pPr")
        indent = OxmlElement("w:ind")
        indent.set(qn("w:left"), "0")
        indent.set(qn("w:firstLine"), "0")
        pPr.append(indent)
        lvl.append(pPr)

        abstract.append(lvl)

    registry._insert_abstract(abstract)

    num_id = registry._next_num_id()
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    ref = OxmlElement("w:abstractNumId")
    ref.set(qn("w:val"), str(abstract_id))
    num.append(ref)
    numbering.append(num)

    for level in range(depth):
        try:
            style = document.styles[f"Heading {level + 1}"]
        except KeyError:
            continue
        pPr = style.element.get_or_add_pPr()
        numPr = get_or_add(pPr, "w:numPr")
        for tag, value in (("w:ilvl", level), ("w:numId", num_id)):
            node = get_or_add(numPr, tag)
            node.set(qn("w:val"), str(value))


# ----------------------------------------------------------------------
# Echte Word-Fussnoten
# ----------------------------------------------------------------------
_FOOTNOTES_SKELETON = (
    XML_DECL
    + f"<w:footnotes {NS_DECL}>"
    + '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr>'
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "<w:r><w:separator/></w:r></w:p></w:footnote>"
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr>'
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
    "</w:footnotes>"
)


class FootnoteStore:
    """Legt den Part word/footnotes.xml an und verwaltet die Fussnoten."""

    def __init__(self, document: Any) -> None:
        self._document = document
        self._part = self._ensure_part()
        self._next_id = self._highest_id() + 1
        self._ensure_settings()

    # -- Aufbau ----------------------------------------------------------
    def _ensure_part(self) -> XmlPart:
        document_part = self._document.part
        for rel in document_part.rels.values():
            if rel.reltype == RT.FOOTNOTES and not rel.is_external:
                return rel.target_part

        element = parse_xml(_FOOTNOTES_SKELETON.encode("utf-8"))
        part = XmlPart(
            PackURI("/word/footnotes.xml"),
            CT.WML_FOOTNOTES,
            element,
            document_part.package,
        )
        document_part.relate_to(part, RT.FOOTNOTES)
        return part

    def _highest_id(self) -> int:
        ids = [
            int(node.get(qn("w:id")))
            for node in self._part.element.findall(qn("w:footnote"))
            if node.get(qn("w:id")) is not None
        ]
        return max(ids, default=0)

    def _ensure_settings(self) -> None:
        """Verweist in settings.xml auf Separator und Fortsetzungstrenner."""
        settings = self._document.settings.element
        if settings.find(qn("w:footnotePr")) is not None:
            return
        fpr = OxmlElement("w:footnotePr")
        for fid in ("-1", "0"):
            ref = OxmlElement("w:footnote")
            ref.set(qn("w:id"), fid)
            fpr.append(ref)
        # w:footnotePr muss vor den meisten anderen Settings stehen
        settings.insert(0, fpr)

    # -- Verwendung ------------------------------------------------------
    def add_footnote(self) -> tuple[int, Any]:
        """Legt eine leere Fussnote an; liefert (id, erstes Absatz-Element)."""
        footnote_id = self._next_id
        self._next_id += 1

        footnote = OxmlElement("w:footnote")
        footnote.set(qn("w:id"), str(footnote_id))

        paragraph = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        style = OxmlElement("w:pStyle")
        style.set(qn("w:val"), "FootnoteText")
        pPr.append(style)
        paragraph.append(pPr)

        marker = OxmlElement("w:r")
        marker_rPr = OxmlElement("w:rPr")
        marker_style = OxmlElement("w:rStyle")
        marker_style.set(qn("w:val"), "FootnoteReference")
        marker_rPr.append(marker_style)
        marker.append(marker_rPr)
        marker.append(OxmlElement("w:footnoteRef"))
        paragraph.append(marker)

        space = OxmlElement("w:r")
        space_text = OxmlElement("w:t")
        space_text.set(qn("xml:space"), "preserve")
        space_text.text = " "
        space.append(space_text)
        paragraph.append(space)

        footnote.append(paragraph)
        self._part.element.append(footnote)
        return footnote_id, paragraph

    def add_paragraph_to(self, footnote_id: int) -> Any:
        """Haengt einen weiteren Absatz an eine bestehende Fussnote an."""
        for node in self._part.element.findall(qn("w:footnote")):
            if node.get(qn("w:id")) == str(footnote_id):
                paragraph = OxmlElement("w:p")
                pPr = OxmlElement("w:pPr")
                style = OxmlElement("w:pStyle")
                style.set(qn("w:val"), "FootnoteText")
                pPr.append(style)
                paragraph.append(pPr)
                node.append(paragraph)
                return paragraph
        raise KeyError(footnote_id)

    @staticmethod
    def add_reference(paragraph: Any, footnote_id: int) -> None:
        """Setzt die hochgestellte Fussnotenziffer in den Fliesstext."""
        run = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        style = OxmlElement("w:rStyle")
        style.set(qn("w:val"), "FootnoteReference")
        rPr.append(style)
        run.append(rPr)

        ref = OxmlElement("w:footnoteReference")
        ref.set(qn("w:id"), str(footnote_id))
        run.append(ref)
        paragraph._p.append(run)

    @property
    def used(self) -> bool:
        return self._next_id > 1


# ----------------------------------------------------------------------
# Tabellen
# ----------------------------------------------------------------------
def set_table_layout_fixed(table: Any) -> None:
    tblPr = table._tbl.tblPr
    layout = get_or_add(tblPr, "w:tblLayout")
    layout.set(qn("w:type"), "fixed")


def set_repeat_header(row: Any) -> None:
    """Wiederholt die Kopfzeile auf jeder Seite."""
    trPr = row._tr.get_or_add_trPr()
    get_or_add(trPr, "w:tblHeader").set(qn("w:val"), "true")


def prevent_row_split(row: Any) -> None:
    trPr = row._tr.get_or_add_trPr()
    get_or_add(trPr, "w:cantSplit").set(qn("w:val"), "true")


def set_cell_margins(table: Any, top: int = 60, start: int = 90, bottom: int = 60, end: int = 90) -> None:
    """Innenabstaende aller Zellen in Twips."""
    tblPr = table._tbl.tblPr
    margins = get_or_add(tblPr, "w:tblCellMar")
    for tag, value in (("w:top", top), ("w:start", start), ("w:bottom", bottom), ("w:end", end)):
        node = get_or_add(margins, tag)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


# ----------------------------------------------------------------------
# Nachbearbeitung
# ----------------------------------------------------------------------
# Geschuetzte Leerzeichen bleiben auch am Absatzrand stehen - wer sie
# schreibt, meint sie.
PROTECTED_SPACES = "\u00a0\u202f\u2007"
_EDGE_SPACE = re.compile(rf"^[^\S{PROTECTED_SPACES}]+|[^\S{PROTECTED_SPACES}]+$")


def trim_paragraph_edges(root: Any, keep_styles: Iterable[str] = ()) -> None:
    """Entfernt fuehrenden und nachlaufenden Leerraum in jedem Absatz.

    HTML laesst Leerraum an Zeilenenden zu, Word zeigt ihn als sichtbare
    Luecke. Codebloecke bleiben ausgenommen, dort ist Einrueckung Inhalt.
    """
    protected = set(keep_styles)
    for paragraph in root.iter(qn("w:p")):
        pPr = paragraph.find(qn("w:pPr"))
        if pPr is not None:
            style = pPr.find(qn("w:pStyle"))
            if style is not None and style.get(qn("w:val")) in protected:
                continue

        texts = [node for node in paragraph.iter(qn("w:t")) if node.text is not None]
        if not texts:
            continue

        first, last = texts[0], texts[-1]
        first.text = _EDGE_SPACE.sub("", first.text) if first is last else re.sub(
            rf"^[^\S{PROTECTED_SPACES}]+", "", first.text
        )
        if first is not last:
            last.text = re.sub(rf"[^\S{PROTECTED_SPACES}]+$", "", last.text)

        for node in (first, last):
            if node.text and (node.text[:1].isspace() or node.text[-1:].isspace()):
                node.set(qn("xml:space"), "preserve")
            elif node.get(qn("xml:space")):
                del node.attrib[qn("xml:space")]


# ----------------------------------------------------------------------
# Sonstiges
# ----------------------------------------------------------------------
def set_document_language(document: Any, lang: str) -> None:
    """Setzt die Bearbeitungssprache im Standard-Zeichenformat."""
    try:
        style = document.styles["Normal"]
    except KeyError:
        return
    rPr = style.element.get_or_add_rPr()
    node = get_or_add(rPr, "w:lang")
    node.set(qn("w:val"), lang)


def set_compat_mode(document: Any) -> None:
    """Markiert das Dokument als Word-2013+-Format (verhindert Kompatibilitaetsmodus)."""
    settings = document.settings.element
    compat = get_or_add(settings, "w:compat")
    setting = OxmlElement("w:compatSetting")
    setting.set(qn("w:name"), "compatibilityMode")
    setting.set(qn("w:uri"), "http://schemas.microsoft.com/office/word")
    setting.set(qn("w:val"), "15")
    compat.append(setting)


def points(value: float) -> Pt:
    return Pt(value)
