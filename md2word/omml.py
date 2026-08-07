"""LaTeX formulas as real Word equations (OMML).

Word stores mathematics as Office Math Markup Language, a dialect of its own
that sits in the `m:` namespace inside the document body. The route taken here
is LaTeX -> MathML (via latex2mathml) -> OMML, with the second step written out
below rather than delegated to Microsoft's MML2OMML.xsl: that stylesheet ships
with Office, is not redistributable, and would tie the converter to a machine
that has Word installed.

Anything this module cannot translate raises :class:`UnsupportedMath`, and the
renderer falls back to formatted text. A formula that comes out slightly plain
is better than a document Word refuses to open.
"""

from __future__ import annotations

from typing import Any, Iterable, Sequence

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

try:
    from latex2mathml.converter import convert as latex_to_mathml

    LATEX_AVAILABLE = True
except ImportError:  # pragma: no cover - optional at runtime
    LATEX_AVAILABLE = False

from lxml import etree

MATHML_NS = "http://www.w3.org/1998/Math/MathML"

# Operators that carry their limits above and below in Word, rather than as
# ordinary sub/superscripts.
N_ARY = {
    "∑",  # n-ary summation
    "∏",  # n-ary product
    "∐",  # n-ary coproduct
    "∫",  # integral
    "∬", "∭",  # double, triple integral
    "∮", "∯", "∰",  # contour integrals
    "⋃", "⋂",  # n-ary union, intersection
    "⋁", "⋀",  # n-ary logical or, and
}

# Characters that sit on top of an expression as an accent instead of being a
# limit above it.
ACCENT_CHARS = {
    "^", "ˆ",          # circumflex
    "~", "˜", "̃",  # tilde
    "¯", "̄", "‾",  # macron / overline
    "˙", "̇",          # dot
    "¨", "̈",          # diaeresis
    "→", "⃗",          # vector arrow
    "ˇ", "̌",          # caron
    "́", "̀",          # acute, grave
}

_BRACKETS = {"(": ")", "[": "]", "{": "}", "⟨": "⟩", "⌊": "⌋", "⌈": "⌉"}
_CLOSERS = {close: open_ for open_, close in _BRACKETS.items()}

# Words that behave like an operator name: upright, and with their subscript
# placed underneath in display style.
_LIMIT_NAMES = {"lim", "max", "min", "sup", "inf", "argmax", "argmin", "limsup", "liminf"}


class UnsupportedMath(Exception):
    """The formula contains something this translator does not handle."""


# ----------------------------------------------------------------------
# Element helpers
# ----------------------------------------------------------------------
def _m(tag: str) -> Any:
    return OxmlElement(f"m:{tag}")


def _val(tag: str, value: str) -> Any:
    node = _m(tag)
    node.set(qn("m:val"), value)
    return node


def _local(node: Any) -> str:
    tag = node.tag
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def _text_of(node: Any) -> str:
    return "".join(node.itertext()).strip()


def _run(text: str, upright: bool = False) -> Any:
    """A single OMML run. Variables stay italic, everything else is upright."""
    run = _m("r")
    if upright:
        rPr = _m("rPr")
        rPr.append(_val("sty", "p"))
        run.append(rPr)
    node = _m("t")
    if text != text.strip():
        node.set(qn("xml:space"), "preserve")
    node.text = text
    run.append(node)
    return run


def _wrap(children: Sequence[Any], tag: str = "e") -> Any:
    """Collects converted nodes into an OMML slot such as m:e, m:num, m:sup."""
    holder = _m(tag)
    for child in children:
        holder.append(child)
    return holder


# ----------------------------------------------------------------------
# MathML -> OMML
# ----------------------------------------------------------------------
def _convert_sequence(nodes: Iterable[Any]) -> list[Any]:
    """Converts a run of sibling nodes, grouping bracket pairs into m:d.

    latex2mathml emits parentheses as plain operators, so `(a+b)` arrives as
    three siblings. Turning a matching pair into a delimiter object is what
    makes Word grow the brackets around a tall fraction instead of leaving
    them at character height.
    """
    items = list(nodes)
    result: list[Any] = []
    index = 0

    while index < len(items):
        node = items[index]
        opener = _bracket_char(node)

        if opener and opener in _BRACKETS:
            closing = _find_closing(items, index, opener)
            if closing is not None:
                inner = _convert_sequence(items[index + 1:closing])
                result.append(_delimiter(opener, _BRACKETS[opener], inner))
                index = closing + 1
                continue

        converted = _convert(node)
        index += 1

        # MathML does not say how far the body of a sum or integral reaches,
        # so the n-ary operator arrives with an empty one. Word would show a
        # placeholder box there, and it is also what a person would type next,
        # so the following item moves inside.
        if index < len(items) and _is_empty_nary(converted):
            follower = _convert(items[index])
            if follower:
                _nary_body(converted[0]).extend(follower)
                index += 1

        result.extend(converted)

    return result


def _is_empty_nary(converted: Sequence[Any]) -> bool:
    if len(converted) != 1 or _local(converted[0]) != "nary":
        return False
    return len(_nary_body(converted[0])) == 0


def _nary_body(nary: Any) -> Any:
    """The m:e slot of an n-ary operator."""
    return nary.find(qn("m:e"))


def _bracket_char(node: Any) -> str | None:
    if _local(node) != "mo":
        return None
    text = _text_of(node)
    return text if len(text) == 1 else None


def _find_closing(items: Sequence[Any], start: int, opener: str) -> int | None:
    """Index of the operator closing the bracket at `start`, or None."""
    wanted = _BRACKETS[opener]
    depth = 0
    for position in range(start + 1, len(items)):
        char = _bracket_char(items[position])
        if char == opener:
            depth += 1
        elif char == wanted:
            if depth == 0:
                return position
            depth -= 1
    return None


def _delimiter(begin: str, end: str, inner: Sequence[Any]) -> Any:
    node = _m("d")
    props = _m("dPr")
    props.append(_val("begChr", begin))
    props.append(_val("endChr", end))
    node.append(props)
    node.append(_wrap(inner))
    return node


def _convert(node: Any) -> list[Any]:
    """Converts one MathML node into zero or more OMML elements."""
    tag = _local(node)
    handler = _HANDLERS.get(tag)
    if handler is None:
        raise UnsupportedMath(f"MathML element <{tag}> is not supported")
    return handler(node)


def _children(node: Any) -> list[Any]:
    return [child for child in node if isinstance(child.tag, str)]


# -- leaves -------------------------------------------------------------
def _h_identifier(node: Any) -> list[Any]:
    text = _text_of(node)
    if not text:
        return []
    # Single letters are variables (italic); longer names are functions such
    # as sin or log and belong upright.
    return [_run(text, upright=len(text) > 1)]


def _h_number(node: Any) -> list[Any]:
    text = _text_of(node)
    return [_run(text, upright=True)] if text else []


def _h_operator(node: Any) -> list[Any]:
    text = _text_of(node)
    return [_run(text, upright=True)] if text else []


def _h_text(node: Any) -> list[Any]:
    text = "".join(node.itertext())
    return [_run(text, upright=True)] if text.strip() else []


def _h_space(node: Any) -> list[Any]:
    return []


# -- containers ---------------------------------------------------------
def _h_row(node: Any) -> list[Any]:
    return _convert_sequence(_children(node))


def _h_fraction(node: Any) -> list[Any]:
    parts = _children(node)
    if len(parts) != 2:
        raise UnsupportedMath("mfrac needs exactly two children")
    frac = _m("f")
    props = _m("fPr")
    props.append(_val("type", "bar"))
    frac.append(props)
    frac.append(_wrap(_convert(parts[0]), "num"))
    frac.append(_wrap(_convert(parts[1]), "den"))
    return [frac]


def _h_sqrt(node: Any) -> list[Any]:
    rad = _m("rad")
    props = _m("radPr")
    props.append(_val("degHide", "1"))
    rad.append(props)
    rad.append(_m("deg"))  # must be present, even when hidden
    rad.append(_wrap(_convert_sequence(_children(node))))
    return [rad]


def _h_root(node: Any) -> list[Any]:
    parts = _children(node)
    if len(parts) != 2:
        raise UnsupportedMath("mroot needs exactly two children")
    rad = _m("rad")
    props = _m("radPr")
    props.append(_val("degHide", "0"))
    rad.append(props)
    rad.append(_wrap(_convert(parts[1]), "deg"))
    rad.append(_wrap(_convert(parts[0])))
    return [rad]


def _base_operator(node: Any) -> str | None:
    """The character of a base that is a bare operator, if it is one."""
    if _local(node) == "mo":
        return _text_of(node)
    if _local(node) == "mrow":
        inner = _children(node)
        if len(inner) == 1 and _local(inner[0]) == "mo":
            return _text_of(inner[0])
    return None


def _nary(char: str, sub: Any | None, sup: Any | None, body: Sequence[Any]) -> Any:
    node = _m("nary")
    props = _m("naryPr")
    props.append(_val("chr", char))
    props.append(_val("limLoc", "undOvr" if char not in "∫∬∭∮" else "subSup"))
    props.append(_val("subHide", "0" if sub is not None else "1"))
    props.append(_val("supHide", "0" if sup is not None else "1"))
    node.append(props)
    node.append(_wrap(_convert(sub) if sub is not None else [], "sub"))
    node.append(_wrap(_convert(sup) if sup is not None else [], "sup"))
    node.append(_wrap(body))
    return [node]


def _h_subsup(node: Any) -> list[Any]:
    parts = _children(node)
    if len(parts) != 3:
        raise UnsupportedMath("msubsup needs exactly three children")
    base, sub, sup = parts

    char = _base_operator(base)
    if char in N_ARY:
        return _nary(char, sub, sup, [])

    element = _m("sSubSup")
    element.append(_wrap(_convert(base)))
    element.append(_wrap(_convert(sub), "sub"))
    element.append(_wrap(_convert(sup), "sup"))
    return [element]


def _h_sub(node: Any) -> list[Any]:
    parts = _children(node)
    if len(parts) != 2:
        raise UnsupportedMath("msub needs exactly two children")
    base, sub = parts

    char = _base_operator(base)
    if char in N_ARY:
        return _nary(char, sub, None, [])
    if char and char.strip().lower() in _LIMIT_NAMES:
        return _limit("limLow", base, sub)

    element = _m("sSub")
    element.append(_wrap(_convert(base)))
    element.append(_wrap(_convert(sub), "sub"))
    return [element]


def _h_sup(node: Any) -> list[Any]:
    parts = _children(node)
    if len(parts) != 2:
        raise UnsupportedMath("msup needs exactly two children")
    base, sup = parts

    char = _base_operator(base)
    if char in N_ARY:
        return _nary(char, None, sup, [])

    element = _m("sSup")
    element.append(_wrap(_convert(base)))
    element.append(_wrap(_convert(sup), "sup"))
    return [element]


def _limit(kind: str, base: Any, limit: Any) -> list[Any]:
    element = _m(kind)
    element.append(_wrap(_convert(base)))
    element.append(_wrap(_convert(limit), "lim"))
    return [element]


def _accent(base: Any, char: str) -> list[Any]:
    element = _m("acc")
    props = _m("accPr")
    props.append(_val("chr", char))
    element.append(props)
    element.append(_wrap(_convert(base)))
    return [element]


def _h_over(node: Any) -> list[Any]:
    parts = _children(node)
    if len(parts) != 2:
        raise UnsupportedMath("mover needs exactly two children")
    base, over = parts

    char = _text_of(over)
    is_accent = node.get("accent") == "true" or char in ACCENT_CHARS
    if is_accent and len(char) == 1:
        return _accent(base, char)
    return _limit("limUpp", base, over)


def _h_under(node: Any) -> list[Any]:
    parts = _children(node)
    if len(parts) != 2:
        raise UnsupportedMath("munder needs exactly two children")
    base, under = parts

    char = _base_operator(base)
    if char in N_ARY:
        return _nary(char, under, None, [])
    return _limit("limLow", base, under)


def _h_underover(node: Any) -> list[Any]:
    parts = _children(node)
    if len(parts) != 3:
        raise UnsupportedMath("munderover needs exactly three children")
    base, under, over = parts

    char = _base_operator(base)
    if char in N_ARY:
        return _nary(char, under, over, [])

    # No direct OMML equivalent: nest the two limits.
    inner = _m("limLow")
    inner.append(_wrap(_convert(base)))
    inner.append(_wrap(_convert(under), "lim"))

    outer = _m("limUpp")
    outer.append(_wrap([inner]))
    outer.append(_wrap(_convert(over), "lim"))
    return [outer]


def _h_table(node: Any) -> list[Any]:
    rows = [child for child in _children(node) if _local(child) == "mtr"]
    if not rows:
        raise UnsupportedMath("mtable without rows")

    columns = max(len(_children(row)) for row in rows)
    matrix = _m("m")
    props = _m("mPr")
    justify = _m("mcJc")
    justify.set(qn("m:val"), "center")
    count = _m("mcs")
    spec = _m("mc")
    spec_pr = _m("mcPr")
    col_count = _m("count")
    col_count.set(qn("m:val"), str(columns))
    spec_pr.append(col_count)
    spec_pr.append(justify)
    spec.append(spec_pr)
    count.append(spec)
    props.append(count)
    matrix.append(props)

    for row in rows:
        mr = _m("mr")
        for cell in _children(row):
            mr.append(_wrap(_convert_sequence(_children(cell))))
        matrix.append(mr)
    return [matrix]


def _h_passthrough(node: Any) -> list[Any]:
    return _convert_sequence(_children(node))


_HANDLERS = {
    "math": _h_passthrough,
    "mrow": _h_row,
    "mi": _h_identifier,
    "mn": _h_number,
    "mo": _h_operator,
    "mtext": _h_text,
    "ms": _h_text,
    "mspace": _h_space,
    "mfrac": _h_fraction,
    "msqrt": _h_sqrt,
    "mroot": _h_root,
    "msub": _h_sub,
    "msup": _h_sup,
    "msubsup": _h_subsup,
    "munder": _h_under,
    "mover": _h_over,
    "munderover": _h_underover,
    "mtable": _h_table,
    "mstyle": _h_passthrough,
    "mpadded": _h_passthrough,
    "menclose": _h_passthrough,
    "merror": _h_passthrough,
    "mphantom": _h_space,
}


# ----------------------------------------------------------------------
# Public API
# ----------------------------------------------------------------------
def mathml_to_omml(mathml: str) -> Any:
    """Converts a MathML string into an m:oMath element."""
    try:
        root = etree.fromstring(mathml.encode("utf-8"))
    except etree.XMLSyntaxError as exc:
        raise UnsupportedMath(f"MathML did not parse: {exc}") from exc

    omath = _m("oMath")
    for child in _convert(root):
        omath.append(child)
    return omath


def latex_to_omml(latex: str) -> Any:
    """Converts a LaTeX formula into an m:oMath element.

    Raises :class:`UnsupportedMath` when the formula cannot be expressed, so
    the caller can fall back to plain text.
    """
    if not LATEX_AVAILABLE:
        raise UnsupportedMath("latex2mathml is not installed")

    source = (latex or "").strip()
    if not source:
        raise UnsupportedMath("empty formula")

    try:
        mathml = latex_to_mathml(source)
    except Exception as exc:  # latex2mathml raises a variety of errors
        raise UnsupportedMath(f"LaTeX did not parse: {exc}") from exc

    return mathml_to_omml(mathml)


def wrap_display(omath: Any, align: str = "center") -> Any:
    """Wraps an equation in m:oMathPara, the block-level form."""
    para = _m("oMathPara")
    props = _m("oMathParaPr")
    props.append(_val("jc", align))
    para.append(props)
    para.append(omath)
    return para
