"""Markdown einlesen: Front Matter abtrennen und nach HTML rendern.

Der Umweg ueber HTML ist Absicht: verschachtelte Strukturen (Listen in
Listen, Zitate mit Codebloecken, Tabellen mit Inline-Auszeichnung) lassen
sich als Baum wesentlich robuster nach Word uebertragen als aus dem
flachen Token-Strom von markdown-it.
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

try:  # PyYAML ist optional - ohne YAML gibt es nur ein Mini-Parsing
    import yaml
except ImportError:  # pragma: no cover
    yaml = None  # type: ignore[assignment]

# Seitenumbruch-Marker, die vor dem Parsen in ein eigenes Tag umgesetzt werden
_PAGEBREAK_PATTERNS = (
    re.compile(r"^[ \t]*<!--\s*(?:pagebreak|page-break|newpage)\s*-->[ \t]*$", re.M | re.I),
    re.compile(r"^[ \t]*\\(?:newpage|pagebreak)[ \t]*$", re.M | re.I),
    re.compile(r"^[ \t]*\{\{\s*(?:pagebreak|newpage)\s*\}\}[ \t]*$", re.M | re.I),
)

_PAGEBREAK_HTML = "\n<div class=\"md2word-pagebreak\"></div>\n"


# Typografische Anfuehrungszeichen je Sprache: oeffnend/schliessend
# aussen, dann innen (einfache Anfuehrung)
NNBSP = "\u202f"  # schmales geschuetztes Leerzeichen

_QUOTES = {
    "de": ("\u201e", "\u201c", "\u201a", "\u2018"),
    "fr": ("\u00ab" + NNBSP, NNBSP + "\u00bb", "\u2039" + NNBSP, NNBSP + "\u203a"),
    "guillemets": ("\u00ab", "\u00bb", "\u2039", "\u203a"),
    "en": ("\u201c", "\u201d", "\u2018", "\u2019"),
}

# Sprachen mit Anfuehrungszeichen unten/oben
_LANGS_DE = ("de", "cs", "sk", "sl", "hr", "hu", "pl", "ro", "lt", "et", "is")
# Franzoesisch setzt Guillemets mit schmalem Abstand, die uebrigen ohne
_LANGS_FR = ("fr",)
_LANGS_GUILLEMETS = ("ru", "es", "it", "pt", "no", "el", "tr", "uk", "be")


def quotes_for(lang: str) -> tuple:
    """Passende Anfuehrungszeichen zur Dokumentsprache.

    Bewusst ein Vierertupel und kein String: markdown-it greift ueber den
    Index zu, sodass ein laengerer String stillschweigend die falschen
    Zeichen liefert - und nur so lassen sich mehrzeichige Eintraege wie
    das schmale geschuetzte Leerzeichen der franzoesischen Typografie
    ueberhaupt ausdruecken.
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
    """Baut die MarkdownIt-Instanz mit allen aktivierten Erweiterungen."""
    md = (
        # html bleibt aktiv, damit rohes HTML als eigener Token erkannt wird -
        # unterdrueckt wird es unten gezielt beim Rendern.
        MarkdownIt(
            "commonmark",
            {
                "html": True,
                "linkify": True,
                "typographer": True,
                "quotes": quotes_for(lang),
            },
        )
        # Das commonmark-Preset schaltet diese Regeln ab - wir wollen sie
        .enable(["table", "strikethrough", "linkify", "replacements", "smartquotes"])
        .use(front_matter_plugin)
        .use(footnote_plugin)
        .use(deflist_plugin)
        .use(tasklists_plugin, enabled=True, label=True)
        .use(dollarmath_plugin, allow_space=True, double_inline=True)
    )

    if not allow_html:
        # Rohes HTML komplett unterdruecken statt es als Text auszugeben
        md.add_render_rule("html_block", lambda *args: "")
        md.add_render_rule("html_inline", lambda *args: "")

    return md


def split_front_matter(text: str) -> tuple[dict[str, Any], str]:
    """Trennt einen YAML-Front-Matter-Block vom eigentlichen Markdown ab."""
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

    # Rueckfallebene: einfache "key: value"-Zeilen
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
    """Konvertiert Markdown nach HTML und liefert zusaetzlich das Front Matter."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    meta, body = split_front_matter(text)
    body = _mark_pagebreaks(body)

    # Die Sprache bestimmt die Anfuehrungszeichen; das Front Matter darf sie
    # setzen, denn es wird vor dem Rendern gelesen.
    lang = str(meta.get("lang") or meta.get("language") or getattr(config, "lang", "en"))
    allow_html = not getattr(config, "strip_html", False)

    md = build_markdown(lang=lang, allow_html=allow_html)
    html = md.render(body)
    return html, meta
