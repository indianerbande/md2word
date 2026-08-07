"""Formulas: LaTeX becomes a real Word equation (OMML), or falls back to text."""

from __future__ import annotations

import zipfile

import pytest
from lxml import etree

from md2word import omml
from md2word.omml import UnsupportedMath, latex_to_omml

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def m(tag: str) -> str:
    return f"{{{M}}}{tag}"


def omml_of(latex: str) -> str:
    """The OMML for a formula, as a string, for structural assertions."""
    return etree.tostring(latex_to_omml(latex), encoding="unicode")


def math_elements(path, tag: str) -> list:
    document = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    return document.findall(f".//{m(tag)}")


# ----------------------------------------------------------------------
# The translator on its own
# ----------------------------------------------------------------------
def test_simple_formula_becomes_omath():
    element = latex_to_omml("x")
    assert element.tag == m("oMath")


@pytest.mark.parametrize(
    "latex, expected_tag",
    [
        (r"\frac{a}{b}", "f"),          # fraction
        (r"\sqrt{x}", "rad"),           # radical
        (r"\sqrt[3]{x}", "rad"),
        (r"x^2", "sSup"),               # superscript
        (r"x_1", "sSub"),               # subscript
        (r"x_1^2", "sSubSup"),
        (r"\sum_{i=1}^{n} i", "nary"),  # n-ary operator
        (r"\int_0^1 x\,dx", "nary"),
        (r"(a+b)", "d"),                # delimiter
        (r"\lim_{x \to 0} f", "limLow"),
        (r"\hat{n}", "acc"),            # accent
        (r"\begin{matrix} a & b \\ c & d \end{matrix}", "m"),
    ],
)
def test_constructs_map_to_their_omml_element(latex, expected_tag):
    assert f"<m:{expected_tag}" in omml_of(latex).replace(
        f'{{{M}}}', "m:"
    ) or latex_to_omml(latex).find(f".//{m(expected_tag)}") is not None


def test_variables_stay_italic_numbers_go_upright():
    """Word italicises maths runs by default; only literals are marked plain."""
    element = latex_to_omml("x2")
    runs = element.findall(f".//{m('r')}")
    styles = []
    for run in runs:
        sty = run.find(f"{m('rPr')}/{m('sty')}")
        styles.append((("".join(run.itertext())), sty.get(m("val")) if sty is not None else None))
    assert ("x", None) in styles, f"variable should stay italic: {styles}"
    assert ("2", "p") in styles, f"number should be upright: {styles}"


def test_function_names_are_upright():
    element = latex_to_omml(r"\sin x")
    runs = {"".join(r.itertext()): r for r in element.findall(f".//{m('r')}")}
    sin = runs["sin"]
    assert sin.find(f"{m('rPr')}/{m('sty')}").get(m("val")) == "p"


def test_integral_limits_sit_beside_the_sign():
    """Sums stack their limits, integrals put them at the side - Word's convention."""
    total = latex_to_omml(r"\sum_{i=1}^{n} i").find(f".//{m('nary')}")
    integral = latex_to_omml(r"\int_0^1 x\,dx").find(f".//{m('nary')}")
    assert total.find(f"{m('naryPr')}/{m('limLoc')}").get(m("val")) == "undOvr"
    assert integral.find(f"{m('naryPr')}/{m('limLoc')}").get(m("val")) == "subSup"


def test_nary_body_is_not_left_empty():
    """An empty body would show as a placeholder box in Word."""
    nary = latex_to_omml(r"\sum_{i=1}^{n} i").find(f".//{m('nary')}")
    body = nary.find(m("e"))
    assert len(body), "the following term should move into the operator body"
    assert "".join(body.itertext()) == "i"


def test_hidden_degree_still_has_the_element():
    """m:rad requires m:deg even when degHide is set."""
    rad = latex_to_omml(r"\sqrt{x}").find(f".//{m('rad')}")
    assert rad.find(f"{m('radPr')}/{m('degHide')}").get(m("val")) == "1"
    assert rad.find(m("deg")) is not None


def test_nested_brackets():
    element = latex_to_omml(r"((a+b))")
    assert len(element.findall(f".//{m('d')}")) == 2


def test_unmatched_bracket_stays_a_character():
    """A stray parenthesis must not swallow the rest of the formula."""
    element = latex_to_omml(r"a)")
    assert element.find(f".//{m('d')}") is None
    assert ")" in "".join(element.itertext())


@pytest.mark.parametrize("latex", ["", "   "])
def test_empty_formula_is_rejected(latex):
    with pytest.raises(UnsupportedMath):
        latex_to_omml(latex)


def test_broken_latex_is_rejected():
    with pytest.raises(UnsupportedMath):
        latex_to_omml(r"\frac{")


def test_unknown_mathml_element_is_rejected():
    with pytest.raises(UnsupportedMath):
        omml.mathml_to_omml(
            '<math xmlns="http://www.w3.org/1998/Math/MathML"><mglyph/></math>'
        )


# ----------------------------------------------------------------------
# End to end, inside the document
# ----------------------------------------------------------------------
def test_inline_formula_lands_in_the_paragraph(convert, assert_valid):
    path, _ = convert("Einstein wrote $E = mc^2$ in 1905.")
    assert_valid(path)
    assert len(math_elements(path, "oMath")) == 1
    assert not math_elements(path, "oMathPara"), "inline maths is not a display block"


def test_display_formula_is_wrapped_and_centred(convert, assert_valid):
    path, _ = convert("$$\n\\frac{a}{b}\n$$")
    assert_valid(path)
    paras = math_elements(path, "oMathPara")
    assert len(paras) == 1
    assert paras[0].find(f"{m('oMathParaPr')}/{m('jc')}").get(m("val")) == "center"


def test_surrounding_text_survives(doc):
    document = doc("Before $x^2$ after.")
    assert "Before" in document.paragraphs[0].text
    assert "after." in document.paragraphs[0].text


def test_math_namespace_is_declared(convert):
    path, _ = convert("$x^2$")
    xml = zipfile.ZipFile(path).read("word/document.xml").decode()
    assert 'xmlns:m="' + M + '"' in xml


def test_unconvertible_formula_falls_back_to_text(convert, doc):
    _path, result = convert(r"$\thisisnotlatex{{{$")
    assert any("formula kept as text" in w for w in result.warnings)
    document = doc(r"$\thisisnotlatex{{{$")
    assert "thisisnotlatex" in document.paragraphs[0].text


def test_fallback_still_produces_a_valid_document(convert, assert_valid):
    path, _ = convert(r"$\frac{$ and $$\begin{broken}$$")
    assert_valid(path)


def test_text_mode_disables_omml(convert, doc):
    path, _ = convert("$E = mc^2$", math_mode="text")
    assert not math_elements(path, "oMath")
    document = doc("$E = mc^2$", math_mode="text")
    assert "E = mc^2" in document.paragraphs[0].text


def test_formula_inside_a_list_and_a_table(convert, assert_valid):
    path, _ = convert(
        "- item with $x^2$\n\n| A |\n|---|\n| $\\frac{a}{b}$ |\n"
    )
    assert_valid(path)
    assert len(math_elements(path, "oMath")) == 2


def test_several_formulas_in_one_paragraph(convert, assert_valid):
    path, _ = convert("$a$ and $b$ and $c$")
    assert_valid(path)
    assert len(math_elements(path, "oMath")) == 3


def test_formula_in_a_heading(convert, assert_valid):
    path, _ = convert("# The $E = mc^2$ chapter")
    assert_valid(path)
    assert len(math_elements(path, "oMath")) == 1


def test_document_with_maths_is_deterministic(tmp_path):
    from md2word.config import Config
    from md2word.converter import convert_text

    source = "$x^2$\n\n$$\\sum_{i=1}^{n} i$$"
    blobs = []
    for index in range(2):
        target = tmp_path / f"m{index}.docx"
        convert_text(source, str(target), Config())
        blobs.append(zipfile.ZipFile(target).read("word/document.xml"))
    assert blobs[0] == blobs[1]
