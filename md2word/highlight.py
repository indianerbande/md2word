"""Syntax highlighting for code blocks via Pygments.

Returns a list of coloured text fragments from which the renderer builds
the individual Word runs.
"""

from __future__ import annotations

from dataclasses import dataclass

try:
    from pygments import lex
    from pygments.lexers import get_lexer_by_name, guess_lexer
    from pygments.styles import get_style_by_name
    from pygments.util import ClassNotFound

    PYGMENTS_AVAILABLE = True
except ImportError:  # pragma: no cover
    PYGMENTS_AVAILABLE = False
    ClassNotFound = Exception  # type: ignore[assignment,misc]


@dataclass(frozen=True)
class Fragment:
    """A contiguous piece of code sharing one set of formatting."""

    text: str
    color: str | None = None
    bold: bool = False
    italic: bool = False
    underline: bool = False


# Aliases common in Markdown but unknown to Pygments
_ALIASES = {
    "js": "javascript",
    "ts": "typescript",
    "sh": "bash",
    "shell": "bash",
    "zsh": "bash",
    "console": "console",
    "py": "python",
    "yml": "yaml",
    "dockerfile": "docker",
    "vue": "html",
    "jsx": "jsx",
    "tsx": "tsx",
    "cs": "csharp",
    "cpp": "cpp",
    "h": "c",
    "md": "markdown",
    "conf": "ini",
    "env": "bash",
    "psl": "powershell",
    "ps1": "powershell",
    "tf": "terraform",
}


def background_for_style(style_name: str) -> str | None:
    """Returns a Pygments style's background colour as hex without the '#'."""
    if not PYGMENTS_AVAILABLE:
        return None
    try:
        style = get_style_by_name(style_name)
    except ClassNotFound:
        return None
    color = getattr(style, "background_color", None)
    if isinstance(color, str) and color.startswith("#") and len(color) == 7:
        return color[1:].upper()
    return None


def _resolve_lexer(language: str, code: str):
    if not language:
        try:
            return guess_lexer(code)
        except Exception:
            return None

    name = _ALIASES.get(language.lower().strip(), language.lower().strip())
    try:
        return get_lexer_by_name(name, stripnl=False, ensurenl=False)
    except ClassNotFound:
        try:
            return get_lexer_by_name(language, stripnl=False, ensurenl=False)
        except ClassNotFound:
            return None


def highlight_code(code: str, language: str = "", style_name: str = "friendly") -> list[Fragment]:
    """Splits source code into coloured fragments.

    Falls back to a single unformatted fragment when Pygments is missing
    or the language cannot be identified.
    """
    if not PYGMENTS_AVAILABLE:
        return [Fragment(code)]

    lexer = _resolve_lexer(language, code)
    if lexer is None:
        return [Fragment(code)]

    try:
        style = get_style_by_name(style_name)
    except ClassNotFound:
        style = get_style_by_name("friendly")

    fragments: list[Fragment] = []
    try:
        tokens = list(lex(code, lexer))
    except Exception:
        return [Fragment(code)]

    for token_type, value in tokens:
        if not value:
            continue
        spec = style.style_for_token(token_type)
        color = spec.get("color")
        fragments.append(
            Fragment(
                text=value,
                color=color.upper() if color else None,
                bold=bool(spec.get("bold")),
                italic=bool(spec.get("italic")),
                underline=bool(spec.get("underline")),
            )
        )

    return fragments or [Fragment(code)]


def available_styles() -> list[str]:
    if not PYGMENTS_AVAILABLE:
        return []
    from pygments.styles import get_all_styles

    return sorted(get_all_styles())


def available_lexers() -> list[str]:
    if not PYGMENTS_AVAILABLE:
        return []
    from pygments.lexers import get_all_lexers

    names: set[str] = set()
    for _long_name, aliases, _patterns, _mimetypes in get_all_lexers():
        names.update(aliases)
    return sorted(names)
