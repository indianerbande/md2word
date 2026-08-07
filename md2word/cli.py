"""Kommandozeilen-Schnittstelle von md2word."""

from __future__ import annotations

import argparse
import glob
import os
import sys
import textwrap
from typing import Any, Sequence

from md2word import __version__
from md2word.config import DEFAULT_THEME, PAGE_SIZES, THEMES, Config

# argparse-Ziel -> Config-Feld (nur wo die Namen abweichen)
_FIELD_MAP = {
    "no_highlight": "highlight",
    "no_remote_images": "download_images",
}

_EPILOG = textwrap.dedent(
    """
    Beispiele:
      md2word bericht.md
      md2word bericht.md -o ~/Desktop/Bericht.docx --toc --page-numbers
      md2word *.md --output-dir build --theme modern --title-page
      md2word handbuch.md --page-size a4 --margin 20 --number-headings --break-on-h1
      cat text.md | md2word - -o ausgabe.docx

    Metadaten und Optionen koennen auch im YAML-Front-Matter stehen:
      ---
      title: Projektbericht
      author: Jane Doe
      date: 2026-08-07
      toc: true
      page_numbers: true
      ---
    """
).strip()


class _Formatter(argparse.RawDescriptionHelpFormatter, argparse.ArgumentDefaultsHelpFormatter):
    """Behaelt die Formatierung des Epilogs und zeigt Standardwerte an."""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="md2word",
        description="Konvertiert Markdown-Dateien in Microsoft-Word-Dokumente (.docx).",
        epilog=_EPILOG,
        formatter_class=_Formatter,
    )

    parser.add_argument(
        "inputs",
        nargs="*",
        metavar="DATEI",
        help="Eine oder mehrere Markdown-Dateien; '-' liest von der Standardeingabe",
    )
    parser.add_argument("-o", "--output", metavar="PFAD", help="Zieldatei (nur bei einer Eingabedatei)")
    parser.add_argument("-d", "--output-dir", metavar="ORDNER", help="Zielordner fuer alle Ausgaben")
    parser.add_argument(
        "-f", "--force", action="store_true", help="Vorhandene Zieldateien ueberschreiben"
    )
    parser.add_argument("-q", "--quiet", action="store_true", help="Keine Statusmeldungen ausgeben")
    parser.add_argument("-v", "--verbose", action="store_true", help="Ausfuehrliche Ausgabe")
    parser.add_argument("--version", action="version", version=f"md2word {__version__}")

    layout = parser.add_argument_group("Seitenlayout")
    layout.add_argument(
        "--page-size", choices=sorted(PAGE_SIZES), default="a4", help="Papierformat"
    )
    layout.add_argument("--landscape", action="store_true", help="Querformat")
    layout.add_argument(
        "--margin", type=float, metavar="MM", help="Alle vier Seitenraender in Millimetern"
    )
    layout.add_argument("--margin-top", type=float, metavar="MM", help="Oberer Rand")
    layout.add_argument("--margin-bottom", type=float, metavar="MM", help="Unterer Rand")
    layout.add_argument("--margin-left", type=float, metavar="MM", help="Linker Rand")
    layout.add_argument("--margin-right", type=float, metavar="MM", help="Rechter Rand")

    typo = parser.add_argument_group("Typografie")
    typo.add_argument(
        "--theme", choices=sorted(THEMES), default=DEFAULT_THEME, help="Farb- und Schriftschema"
    )
    typo.add_argument("--body-font", metavar="NAME", help="Schriftart des Fliesstextes")
    typo.add_argument("--heading-font", metavar="NAME", help="Schriftart der Ueberschriften")
    typo.add_argument("--code-font", metavar="NAME", help="Schriftart fuer Quelltext")
    typo.add_argument("--font-size", type=float, metavar="PT", help="Grundschriftgroesse in Punkt")
    typo.add_argument("--code-font-size", type=float, metavar="PT", help="Schriftgroesse im Code")
    typo.add_argument("--line-spacing", type=float, metavar="FAKTOR", help="Zeilenabstand")
    typo.add_argument("--accent", metavar="HEX", help="Akzentfarbe, z. B. 2F5496")
    typo.add_argument("--link-color", metavar="HEX", help="Farbe fuer Hyperlinks")
    typo.add_argument("--code-bg", metavar="HEX", help="Hintergrundfarbe von Codebloecken")

    structure = parser.add_argument_group("Dokumentstruktur")
    structure.add_argument("--toc", action="store_true", help="Inhaltsverzeichnis voranstellen")
    structure.add_argument(
        "--toc-depth", type=int, default=3, metavar="N", help="Ueberschriftenebenen im Verzeichnis"
    )
    structure.add_argument("--toc-title", metavar="TEXT", help="Ueberschrift des Inhaltsverzeichnisses")
    structure.add_argument(
        "--title-page", action="store_true", help="Eigene Titelseite aus den Metadaten erzeugen"
    )
    structure.add_argument(
        "--number-headings", action="store_true", help="Ueberschriften automatisch nummerieren"
    )
    structure.add_argument(
        "--page-numbers", action="store_true", help="Seitenzahlen in die Fusszeile setzen"
    )
    structure.add_argument("--header-text", metavar="TEXT", help="Text der Kopfzeile")
    structure.add_argument("--footer-text", metavar="TEXT", help="Text der Fusszeile")
    structure.add_argument(
        "--break-on-h1", action="store_true", help="Vor jeder H1-Ueberschrift eine neue Seite"
    )

    content = parser.add_argument_group("Inhalte")
    content.add_argument(
        "--no-highlight", action="store_true", help="Syntaxhervorhebung abschalten"
    )
    content.add_argument(
        "--pygments-style", default="friendly", metavar="NAME", help="Farbschema fuer Quelltext"
    )
    content.add_argument(
        "--list-pygments-styles",
        action="store_true",
        help="Verfuegbare Syntax-Farbschemata auflisten und beenden",
    )
    content.add_argument(
        "--no-remote-images", action="store_true", help="Bilder von http(s)-URLs nicht laden"
    )
    content.add_argument(
        "--max-image-width", type=float, metavar="MM", help="Maximale Bildbreite in Millimetern"
    )
    content.add_argument(
        "--captions",
        choices=("title", "alt", "none"),
        default="title",
        help="Woraus Bildunterschriften entstehen: aus dem Titel in Anfuehrungszeichen, "
        "ersatzweise aus dem Alternativtext, oder gar nicht",
    )
    content.add_argument(
        "--footnotes",
        dest="footnote_mode",
        choices=("footnotes", "endnotes"),
        default="footnotes",
        help="Echte Word-Fussnoten oder gesammelte Anmerkungen am Dokumentende",
    )
    content.add_argument(
        "--strip-html",
        action="store_true",
        help="Rohes HTML im Markdown ignorieren statt es zu uebernehmen",
    )
    content.add_argument("--lang", default="de-DE", metavar="CODE", help="Sprache des Dokuments")

    meta = parser.add_argument_group("Metadaten (ueberschreiben das Front Matter)")
    meta.add_argument("--title", metavar="TEXT", help="Dokumenttitel")
    meta.add_argument("--subtitle", metavar="TEXT", help="Untertitel")
    meta.add_argument("--author", metavar="TEXT", help="Verfasser")
    meta.add_argument("--date", metavar="TEXT", help="Datumsangabe")
    meta.add_argument("--subject", metavar="TEXT", help="Thema")
    meta.add_argument("--keywords", metavar="TEXT", help="Schlagworte, kommagetrennt")

    advanced = parser.add_argument_group("Erweitert")
    advanced.add_argument(
        "--reference-doc",
        metavar="DATEI",
        help="Bestehendes .docx als Formatvorlage verwenden (Inhalt wird geleert)",
    )

    return parser


# ----------------------------------------------------------------------
def _explicit_options(argv: Sequence[str], parser: argparse.ArgumentParser) -> set[str]:
    """Ermittelt, welche Config-Felder wirklich auf der Kommandozeile standen."""
    seen: set[str] = set()
    long_options: dict[str, str] = {}
    for action in parser._actions:
        for option in action.option_strings:
            long_options[option] = action.dest

    for token in argv:
        name = token.split("=", 1)[0]
        dest = long_options.get(name)
        if dest:
            seen.add(_FIELD_MAP.get(dest, dest))
    return seen


def config_from_args(args: argparse.Namespace, explicit: set[str]) -> Config:
    """Baut die Konfiguration aus den geparsten Argumenten."""
    values: dict[str, Any] = {
        "page_size": args.page_size,
        "landscape": args.landscape,
        "theme": args.theme,
        "toc": args.toc,
        "toc_depth": args.toc_depth,
        "title_page": args.title_page,
        "number_headings": args.number_headings,
        "page_numbers": args.page_numbers,
        "break_on_h1": args.break_on_h1,
        "highlight": not args.no_highlight,
        "pygments_style": args.pygments_style,
        "download_images": not args.no_remote_images,
        "strip_html": args.strip_html,
        "captions": args.captions,
        "footnote_mode": args.footnote_mode,
        "lang": args.lang,
        "verbose": args.verbose,
    }

    optional = (
        "toc_title", "header_text", "footer_text", "body_font", "heading_font",
        "code_font", "font_size", "code_font_size", "line_spacing", "accent",
        "link_color", "code_bg", "max_image_width", "title", "subtitle",
        "author", "date", "subject", "keywords", "reference_doc",
    )
    for name in optional:
        value = getattr(args, name, None)
        if value not in (None, ""):
            values[name] = value

    if args.margin is not None:
        for side in ("margin_top", "margin_bottom", "margin_left", "margin_right"):
            values[side] = args.margin
            explicit.add(side)
    for side in ("margin_top", "margin_bottom", "margin_left", "margin_right"):
        value = getattr(args, side, None)
        if value is not None:
            values[side] = value

    # Passt der Codeblock-Hintergrund nicht zum Farbschema, aus Pygments ableiten
    if "code_bg" not in values and "pygments_style" in explicit:
        from md2word.highlight import background_for_style

        derived = background_for_style(args.pygments_style)
        if derived:
            values["code_bg"] = derived

    config = Config(**values)
    config._explicit = explicit
    return config


def _expand_inputs(patterns: Sequence[str]) -> list[str]:
    """Loest Glob-Muster auf; die Shell erledigt das meist schon."""
    result: list[str] = []
    for pattern in patterns:
        if pattern == "-":
            result.append(pattern)
            continue
        if any(char in pattern for char in "*?[") and not os.path.exists(pattern):
            matches = sorted(glob.glob(pattern, recursive=True))
            if not matches:
                raise FileNotFoundError(pattern)
            result.extend(matches)
        else:
            result.append(pattern)
    return result


def _output_path(input_path: str, args: argparse.Namespace) -> str:
    if args.output and input_path != "-":
        return args.output
    if args.output and input_path == "-":
        return args.output

    stem = "dokument" if input_path == "-" else os.path.splitext(os.path.basename(input_path))[0]
    filename = f"{stem}.docx"

    if args.output_dir:
        return os.path.join(args.output_dir, filename)
    if input_path == "-":
        return os.path.join(os.getcwd(), filename)
    return os.path.join(os.path.dirname(os.path.abspath(input_path)), filename)


# ----------------------------------------------------------------------
def main(argv: Sequence[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.list_pygments_styles:
        from md2word.highlight import available_styles

        for name in available_styles():
            print(name)
        return 0

    if not args.inputs:
        parser.print_help()
        return 2

    if args.output and len([p for p in args.inputs if p != "-"]) > 1:
        parser.error("--output funktioniert nur mit genau einer Eingabedatei; nutze --output-dir")

    try:
        inputs = _expand_inputs(args.inputs)
    except FileNotFoundError as exc:
        print(f"Fehler: keine Datei passt auf das Muster '{exc}'", file=sys.stderr)
        return 1

    if args.output and len(inputs) > 1:
        parser.error("--output funktioniert nur mit genau einer Eingabedatei; nutze --output-dir")

    explicit = _explicit_options(argv, parser)
    config = config_from_args(args, explicit)

    from md2word.converter import convert_file

    failures = 0
    for input_path in inputs:
        if input_path != "-" and not os.path.isfile(input_path):
            print(f"Fehler: Datei nicht gefunden: {input_path}", file=sys.stderr)
            failures += 1
            continue

        output_path = _output_path(input_path, args)

        if os.path.exists(output_path) and not args.force:
            if os.path.abspath(output_path) == os.path.abspath(input_path):
                print(f"Fehler: Ziel und Quelle sind identisch: {output_path}", file=sys.stderr)
                failures += 1
                continue
            print(
                f"Uebersprungen (existiert bereits, --force zum Ueberschreiben): {output_path}",
                file=sys.stderr,
            )
            failures += 1
            continue

        try:
            result = convert_file(input_path, output_path, config)
        except FileNotFoundError as exc:
            print(f"Fehler: {exc}", file=sys.stderr)
            failures += 1
            continue
        except Exception as exc:
            print(f"Fehler bei '{input_path}': {exc}", file=sys.stderr)
            if args.verbose:
                import traceback

                traceback.print_exc()
            failures += 1
            continue

        if not args.quiet:
            source = "Standardeingabe" if input_path == "-" else input_path
            size = os.path.getsize(result.output_path) / 1024
            print(f"{source} -> {result.output_path}  ({size:.0f} KB, {result.heading_count} Ueberschriften)")
        for warning in result.warnings:
            print(f"  Hinweis: {warning}", file=sys.stderr)

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
