"""Character formatting, links, footnotes, images and raw HTML."""

from __future__ import annotations

import base64

import pytest

from md2word import styles as st
from tests.conftest import find_paragraph, w

# 1x1 pixel PNG
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def runs_of(document, needle: str):
    return find_paragraph(document, needle).runs


# ----------------------------------------------------------------------
# Character formatting
# ----------------------------------------------------------------------
@pytest.mark.parametrize(
    "markdown, attribute",
    [
        ("**bold**", "bold"),
        ("*italic*", "italic"),
        ("~~gone~~", "strike"),
    ],
)
def test_basic_emphasis(doc, markdown, attribute):
    document = doc(f"Text with {markdown} in it.")
    marked = [
        r
        for r in document.paragraphs[0].runs
        if (r.font.strike if attribute == "strike" else getattr(r, attribute))
    ]
    assert len(marked) == 1
    assert marked[0].text in {"bold", "italic", "gone"}


def test_nested_emphasis(doc):
    document = doc("***bold and italic***")
    run = [r for r in document.paragraphs[0].runs if r.text == "bold and italic"][0]
    assert run.bold and run.italic


def test_inline_code_style(doc):
    document = doc("A `snippet` in the text.")
    code = [r for r in document.paragraphs[0].runs if r.text == "snippet"][0]
    rStyle = code._r.find(w("rPr")).find(w("rStyle"))
    assert rStyle.get(w("val")) == st.S_CODE_INLINE


def test_inline_code_keeps_spaces(doc):
    document = doc("Before `a  b` after")
    assert "a  b" in [r.text for r in document.paragraphs[0].runs]


def test_superscript_and_subscript(doc):
    document = doc("H<sub>2</sub>O and x<sup>3</sup>")
    aligns = {
        r.text: r._r.find(w("rPr")).find(w("vertAlign")).get(w("val"))
        for r in document.paragraphs[0].runs
        if r._r.find(w("rPr")) is not None
        and r._r.find(w("rPr")).find(w("vertAlign")) is not None
    }
    assert aligns == {"2": "subscript", "3": "superscript"}


def test_raw_html_bold_and_color(doc):
    document = doc('Text <b>bold</b> and <span style="color:#C00000">red</span>.')
    runs = {r.text: r for r in document.paragraphs[0].runs}
    assert runs["bold"].bold
    assert str(runs["red"].font.color.rgb) == "C00000"


def test_strip_html_option(doc):
    document = doc("Paragraph\n\n<div>dropped</div>", strip_html=True)
    assert "dropped" not in " ".join(p.text for p in document.paragraphs)


def test_html_kept_by_default(doc):
    document = doc("Paragraph\n\n<div>kept</div>")
    assert "kept" in " ".join(p.text for p in document.paragraphs)


def test_typographic_quotes(doc):
    document = doc('She said "Hello" -- and left.')
    text = document.paragraphs[0].text
    assert "“" in text or "„" in text
    assert "–" in text, "-- becomes an en dash"


def test_umlauts_and_symbols_survive(doc):
    document = doc("Äöü ÄÖÜ ß € → ✓ 😀")
    assert document.paragraphs[0].text == "Äöü ÄÖÜ ß € → ✓ 😀"


def test_nbsp_is_not_collapsed(doc):
    """A non-breaking space must not degrade into an ordinary one.

    Otherwise the break protection the author intended is lost.
    """
    document = doc("Weight: 10&nbsp;kg")
    assert "10 kg" in document.paragraphs[0].text


def test_nbsp_at_paragraph_end_survives_trimming(doc):
    document = doc("Line ends with protection:&nbsp;")
    assert document.paragraphs[0].text.endswith(" ")


def test_ordinary_trailing_space_is_removed(doc):
    document = doc("Text with trailing space   \n\nSecond paragraph")
    assert not document.paragraphs[0].text.endswith(" ")


# ----------------------------------------------------------------------
# Links
# ----------------------------------------------------------------------
def test_external_link_creates_relationship(convert):
    import zipfile

    from lxml import etree

    path, _ = convert("[Target](https://example.com/path)")
    rels = etree.fromstring(zipfile.ZipFile(path).read("word/_rels/document.xml.rels"))
    targets = [
        r.get("Target")
        for r in rels
        if r.get("TargetMode") == "External"
    ]
    assert "https://example.com/path" in targets


def test_link_text_is_inside_hyperlink_element(doc):
    document = doc("[Click me](https://example.com)")
    paragraph = document.paragraphs[0]
    hyperlink = paragraph._p.find(w("hyperlink"))
    assert hyperlink is not None
    assert "".join(t.text or "" for t in hyperlink.iter(w("t"))) == "Click me"


def test_internal_link_to_heading(doc):
    document = doc("# My Target\n\nSee [there](#my-target).")
    paragraph = find_paragraph(document, "See")
    hyperlink = paragraph._p.find(w("hyperlink"))
    assert hyperlink is not None
    assert hyperlink.get(w("anchor")) == "my_target"


def test_dangling_internal_link_warns(convert):
    _path, result = convert("See [nowhere](#does-not-exist).")
    assert any("internal link" in msg for msg in result.warnings)


def test_autolink(doc):
    document = doc("See https://example.com for more.")
    assert document.paragraphs[0]._p.find(w("hyperlink")) is not None


def test_formatted_link_text(doc):
    document = doc("[**bold** link](https://example.com)")
    hyperlink = document.paragraphs[0]._p.find(w("hyperlink"))
    bolds = [
        r
        for r in hyperlink.iter(w("r"))
        if r.find(w("rPr")) is not None and r.find(w("rPr")).find(w("b")) is not None
    ]
    assert bolds


# ----------------------------------------------------------------------
# Footnotes
# ----------------------------------------------------------------------
def test_real_footnote_creates_part(convert, assert_valid):
    import zipfile

    path, _ = convert("Text[^1]\n\n[^1]: The note.")
    assert_valid(path)
    names = zipfile.ZipFile(path).namelist()
    assert "word/footnotes.xml" in names


def test_footnote_content(convert):
    import zipfile

    from lxml import etree

    path, _ = convert("Text[^a]\n\n[^a]: Body of the **footnote**.")
    footnotes = etree.fromstring(zipfile.ZipFile(path).read("word/footnotes.xml"))
    body = " ".join(t.text or "" for t in footnotes.iter(w("t")))
    assert "Body of the" in body and "footnote" in body


def test_footnote_reference_is_superscript_style(doc):
    document = doc("Text[^1]\n\n[^1]: Note")
    reference = list(document.paragraphs[0]._p.iter(w("footnoteReference")))
    assert len(reference) == 1


def test_footnote_backref_arrow_removed(convert):
    import zipfile

    from lxml import etree

    path, _ = convert("Text[^1]\n\n[^1]: Note")
    footnotes = etree.fromstring(zipfile.ZipFile(path).read("word/footnotes.xml"))
    body = "".join(t.text or "" for t in footnotes.iter(w("t")))
    assert "↩" not in body


def test_endnote_mode(convert, assert_valid):
    import zipfile

    path, _ = convert("Text[^1]\n\n[^1]: The note.", footnote_mode="endnotes")
    assert_valid(path)
    assert "word/footnotes.xml" not in zipfile.ZipFile(path).namelist()

    from docx import Document

    document = Document(str(path))
    assert any(p.text == "Notes" for p in document.paragraphs)
    assert any("The note." in p.text for p in document.paragraphs)


def test_multiple_footnotes_numbered(convert):
    import zipfile

    from lxml import etree

    path, _ = convert("A[^1] B[^2]\n\n[^1]: one\n[^2]: two")
    footnotes = etree.fromstring(zipfile.ZipFile(path).read("word/footnotes.xml"))
    ids = sorted(
        int(n.get(w("id")))
        for n in footnotes.findall(w("footnote"))
        if int(n.get(w("id"))) > 0
    )
    assert ids == [1, 2]


# ----------------------------------------------------------------------
# Images
# ----------------------------------------------------------------------
def test_local_image_embedded(tmp_path, convert, assert_valid):
    import zipfile

    (tmp_path / "bild.png").write_bytes(PNG_1PX)
    path, _ = convert("![](bild.png)")
    assert_valid(path)
    media = [n for n in zipfile.ZipFile(path).namelist() if n.startswith("word/media/")]
    assert len(media) == 1


def test_missing_image_warns_and_continues(convert):
    _path, result = convert("![Alt text](missing.png)")
    assert any("image" in msg for msg in result.warnings)


def test_missing_image_leaves_placeholder(doc):
    document = doc("![Alt text](missing.png)")
    assert "[Image: Alt text]" in document.paragraphs[0].text


def test_data_uri_image(convert, assert_valid):
    import zipfile

    encoded = base64.b64encode(PNG_1PX).decode()
    path, _ = convert(f"![](data:image/png;base64,{encoded})")
    assert_valid(path)
    media = [n for n in zipfile.ZipFile(path).namelist() if n.startswith("word/media/")]
    assert len(media) == 1


def test_remote_images_can_be_disabled(convert):
    _path, result = convert("![](https://example.invalid/x.png)", download_images=False)
    assert any("skipped" in msg for msg in result.warnings)


def test_image_caption_from_title(tmp_path, doc):
    (tmp_path / "b.png").write_bytes(PNG_1PX)
    document = doc('![alt](b.png "Figure 1")')
    captions = [p for p in document.paragraphs if p.style.style_id == st.S_CAPTION]
    assert [p.text for p in captions] == ["Figure 1"]


def test_alt_text_is_no_caption_by_default(tmp_path, doc):
    (tmp_path / "b.png").write_bytes(PNG_1PX)
    document = doc("![just alt](b.png)")
    captions = [p for p in document.paragraphs if p.style.style_id == st.S_CAPTION]
    assert not captions


def test_caption_mode_alt(tmp_path, doc):
    (tmp_path / "b.png").write_bytes(PNG_1PX)
    document = doc("![just alt](b.png)", captions="alt")
    captions = [p for p in document.paragraphs if p.style.style_id == st.S_CAPTION]
    assert [p.text for p in captions] == ["just alt"]


def test_image_scaled_to_text_width(tmp_path, doc):
    from PIL import Image

    Image.new("RGB", (2000, 1000), "blue").save(tmp_path / "large.png", dpi=(96, 96))
    document = doc("![](large.png)")
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    extents = [
        int(node.get("cx"))
        for p in document.paragraphs
        for node in p._p.iter(A + "ext")
        if node.get("cx")
    ]
    assert extents, "no image found"
    # A4 with 25 mm margins -> 160 mm of text width
    assert abs(extents[0] / 36000 - 160.0) < 1.0


def test_inline_image_stays_in_paragraph(tmp_path, doc):
    (tmp_path / "b.png").write_bytes(PNG_1PX)
    document = doc("Before ![](b.png) after")
    paragraph = find_paragraph(document, "Before")
    assert "after" in paragraph.text


# Formulas now become real Word equations; see tests/test_math.py.
