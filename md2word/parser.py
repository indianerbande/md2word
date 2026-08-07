"""Reading Markdown: split off the front matter and render to HTML.

Routing through HTML is deliberate: nested structures (lists within lists,
quotes containing code blocks, tables with inline markup) transfer to Word
far more reliably from a tree than from markdown-it's flat token stream.
"""

from __future__ import annotations

import re
from typing import Any

from markdown_it import MarkdownIt
from mdit_py_plugins.deflist import deflist_plugin
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.footnote import footnote_plugin
from mdit_py_plugins.front_matter import front_matter_plugin
from mdit_py_plugins.tasklists import tasklists_plugin

try:  # PyYAML is optional - without it there is only minimal parsing
    import yaml
except ImportError:  # pragma: no cover
    yaml = None  # type: ignore[assignment]

# Page-break markers, rewritten into a tag of our own before parsing
_PAGEBREAK_PATTERNS = (
    re.compile(r"^[ \t]*<!--\s*(?:pagebreak|page-break|newpage)\s*-->[ \t]*$", re.M | re.I),
    re.compile(r"^[ \t]*\\(?:newpage|pagebreak)[ \t]*$", re.M | re.I),
    re.compile(r"^[ \t]*\{\{\s*(?:pagebreak|newpage)\s*\}\}[ \t]*$", re.M | re.I),
)

_PAGEBREAK_HTML = "\n<div class=\"md2word-pagebreak\"></div>\n"


# Typographic quotation marks per language: opening and closing for the
# outer level, then for the inner one.
NNBSP = "\u202f"  # narrow no-break space

_QUOTES = {
    "de": ("\u201e", "\u201c", "\u201a", "\u2018"),
    "fr": ("\u00ab" + NNBSP, NNBSP + "\u00bb", "\u2039" + NNBSP, NNBSP + "\u203a"),
    "guillemets": ("\u00ab", "\u00bb", "\u2039", "\u203a"),
    "en": ("\u201c", "\u201d", "\u2018", "\u2019"),
}

# Languages using low-then-high quotation marks
_LANGS_DE = ("de", "cs", "sk", "sl", "hr", "hu", "pl", "ro", "lt", "et", "is")
# French sets guillemets with a narrow space, the others without
_LANGS_FR = ("fr",)
_LANGS_GUILLEMETS = ("ru", "es", "it", "pt", "no", "el", "tr", "uk", "be")


def quotes_for(lang: str) -> tuple:
    """The quotation marks matching the document language.

    Deliberately a four-tuple and not a string: markdown-it indexes into
    the value, so an over-long string silently supplies the wrong
    characters - and only a tuple can express multi-character entries such
    as the narrow no-break space of French typography.
    """
    prefix = (lang or "en").split("-")[0].split("_")[0].lower()
    if prefix in _LANGS_DE:
        return _QUOTES["de"]
    if prefix in _LANGS_FR:
        return _QUOTES["fr"]
    if prefix in _LANGS_GUILLEMETS:
        return _QUOTES["guillemets"]
    return _QUOTES["en"]


def build_markdown(lang: str = "en", allow_html: bool = True) -> MarkdownIt:
    """Builds the MarkdownIt instance with every extension enabled."""
    md = (
        # html stays on so raw HTML becomes its own token - suppressing it
        # happens further down, at render time.
        MarkdownIt(
            "commonmark",
            {
                "html": True,
                "linkify": True,
                "typographer": True,
                "quotes": quotes_for(lang),
            },
        )
        # The commonmark preset disables these rules - we want them
        .enable(["table", "strikethrough", "linkify", "replacements", "smartquotes"])
        .use(front_matter_plugin)
        .use(footnote_plugin)
        .use(deflist_plugin)
        .use(tasklists_plugin, enabled=True, label=True)
        # allow_space=False on purpose: with spaces allowed, "costs $100 and
        # $50" reads as one formula spanning both amounts. Requiring the
        # delimiters to hug their content is also what MathJax does.
        .use(dollarmath_plugin, allow_space=False, double_inline=True)
    )

    if not allow_html:
        # Drop raw HTML entirely instead of emitting it as text
        md.add_render_rule("html_block", lambda *args: "")
        md.add_render_rule("html_inline", lambda *args: "")

    return md


def split_front_matter(text: str) -> tuple[dict[str, Any], str]:
    """Separates a YAML front-matter block from the Markdown itself."""
    if not text.startswith("---"):
        return {}, text

    match = re.match(r"^---[ \t]*\r?\n(.*?)\r?\n(?:---|\.\.\.)[ \t]*(?:\r?\n|$)", text, re.S)
    if not match:
        return {}, text

    raw, rest = match.group(1), text[match.end():]
    data = _parse_yaml(raw)
    return data, rest


def _parse_yaml(raw: str) -> dict[str, Any]:
    if yaml is not None:
        try:
            data = yaml.safe_load(raw)
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    # Fallback: plain "key: value" lines
    data: dict[str, Any] = {}
    for line in raw.splitlines():
        if ":" in line and not line.lstrip().startswith("#"):
            key, _, value = line.partition(":")
            data[key.strip()] = value.strip().strip("\"'")
    return data


def _mark_pagebreaks(text: str) -> str:
    for pattern in _PAGEBREAK_PATTERNS:
        text = pattern.sub(_PAGEBREAK_HTML, text)
    return text


def markdown_to_html(text: str, config: Any = None) -> tuple[str, dict[str, Any]]:
    """Converts Markdown to HTML and returns the front matter alongside."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    meta, body = split_front_matter(text)
    body = _mark_pagebreaks(body)

    # The language decides the quotation marks; the front matter may set it,
    # because it is read before rendering starts.
    lang = str(meta.get("lang") or meta.get("language") or getattr(config, "lang", "en"))
    allow_html = not getattr(config, "strip_html", False)

    md = build_markdown(lang=lang, allow_html=allow_html)
    html = md.render(body)
    return html, meta
