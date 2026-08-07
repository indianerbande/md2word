"""Command-line interface of md2word."""

from __future__ import annotations

import argparse
import glob
import os
import sys
import textwrap
from typing import Any, Sequence

from md2word import __version__
from md2word.config import DEFAULT_THEME, PAGE_SIZES, THEMES, Config

# argparse destination -> Config field (only where the names differ)
_FIELD_MAP = {
    "no_highlight": "highlight",
    "no_remote_images": "download_images",
}

_EPILOG = textwrap.dedent(
    """
    Examples:
      md2word report.md
      md2word report.md -o ~/Desktop/Report.docx --toc --page-numbers
      md2word *.md --output-dir build --theme modern --title-page
      md2word manual.md --page-size a4 --margin 20 --number-headings --break-on-h1
      cat notes.md | md2word - -o output.docx

    Metadata and layout options can also live in the YAML front matter:
      ---
      title: Quarterly Report
      author: Jane Doe
      date: 2026-08-07
      toc: true
      page_numbers: true
      ---
    """
).strip()


class _Formatter(argparse.RawDescriptionHelpFormatter, argparse.ArgumentDefaultsHelpFormatter):
    """Keeps the epilog's formatting and shows default values."""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="md2word",
        description="Convert Markdown files into Microsoft Word documents (.docx).",
        epilog=_EPILOG,
        formatter_class=_Formatter,
    )

    parser.add_argument(
        "inputs",
        nargs="*",
        metavar="FILE",
        help="One or more Markdown files; '-' reads from standard input",
    )
    parser.add_argument(
        "-o", "--output", metavar="PATH", help="Target file (only with a single input file)"
    )
    parser.add_argument(
        "-d", "--output-dir", metavar="DIR", help="Target directory for all output"
    )
    parser.add_argument(
        "-f", "--force", action="store_true", help="Overwrite existing target files"
    )
    parser.add_argument("-q", "--quiet", action="store_true", help="Suppress status messages")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")
    parser.add_argument("--version", action="version", version=f"md2word {__version__}")

    layout = parser.add_argument_group("Page layout")
    layout.add_argument(
        "--page-size", choices=sorted(PAGE_SIZES), default="a4", help="Paper size"
    )
    layout.add_argument("--landscape", action="store_true", help="Landscape orientation")
    layout.add_argument(
        "--margin", type=float, metavar="MM", help="All four page margins, in millimetres"
    )
    layout.add_argument("--margin-top", type=float, metavar="MM", help="Top margin")
    layout.add_argument("--margin-bottom", type=float, metavar="MM", help="Bottom margin")
    layout.add_argument("--margin-left", type=float, metavar="MM", help="Left margin")
    layout.add_argument("--margin-right", type=float, metavar="MM", help="Right margin")

    typo = parser.add_argument_group("Typography")
    typo.add_argument(
        "--theme", choices=sorted(THEMES), default=DEFAULT_THEME, help="Colour and font scheme"
    )
    typo.add_argument("--body-font", metavar="NAME", help="Body text font")
    typo.add_argument("--heading-font", metavar="NAME", help="Heading font")
    typo.add_argument("--code-font", metavar="NAME", help="Font for source code")
    typo.add_argument("--font-size", type=float, metavar="PT", help="Base font size, in points")
    typo.add_argument("--code-font-size", type=float, metavar="PT", help="Font size inside code")
    typo.add_argument("--line-spacing", type=float, metavar="FACTOR", help="Line spacing")
    typo.add_argument("--accent", metavar="HEX", help="Accent colour, e.g. 2F5496")
    typo.add_argument("--link-color", metavar="HEX", help="Hyperlink colour")
    typo.add_argument("--code-bg", metavar="HEX", help="Background colour of code blocks")

    structure = parser.add_argument_group("Document structure")
    structure.add_argument("--toc", action="store_true", help="Prepend a table of contents")
    structure.add_argument(
        "--toc-depth", type=int, default=3, metavar="N", help="Heading levels to include"
    )
    structure.add_argument(
        "--toc-title",
        metavar="TEXT",
        help="Heading above the table of contents (default: follows --lang)",
    )
    structure.add_argument(
        "--title-page", action="store_true", help="Build a title page from the metadata"
    )
    structure.add_argument(
        "--number-headings", action="store_true", help="Number headings automatically"
    )
    structure.add_argument(
        "--page-numbers", action="store_true", help="Put page numbers in the footer"
    )
    structure.add_argument("--header-text", metavar="TEXT", help="Header text")
    structure.add_argument("--footer-text", metavar="TEXT", help="Footer text")
    structure.add_argument(
        "--break-on-h1", action="store_true", help="Page break before every level-1 heading"
    )

    content = parser.add_argument_group("Content")
    content.add_argument(
        "--no-highlight", action="store_true", help="Turn off syntax highlighting"
    )
    content.add_argument(
        "--pygments-style", default="friendly", metavar="NAME", help="Colour scheme for code"
    )
    content.add_argument(
        "--list-pygments-styles",
        action="store_true",
        help="List the available syntax colour schemes and exit",
    )
    content.add_argument(
        "--no-remote-images", action="store_true", help="Do not fetch images from http(s) URLs"
    )
    content.add_argument(
        "--max-image-width", type=float, metavar="MM", help="Maximum image width, in millimetres"
    )
    content.add_argument(
        "--captions",
        choices=("title", "alt", "none"),
        default="title",
        help="What image captions are made from: the quoted title, the alt text "
        "as a fallback, or nothing",
    )
    content.add_argument(
        "--footnotes",
        dest="footnote_mode",
        choices=("footnotes", "endnotes"),
        default="footnotes",
        help="Real Word footnotes, or notes collected at the end of the document",
    )
    content.add_argument(
        "--strip-html",
        action="store_true",
        help="Ignore raw HTML instead of carrying it through",
    )
    content.add_argument(
        "--lang",
        default="en-US",
        metavar="CODE",
        help="Document language; also selects quotation marks and built-in headings",
    )

    meta = parser.add_argument_group("Metadata (overrides the front matter)")
    meta.add_argument("--title", metavar="TEXT", help="Document title")
    meta.add_argument("--subtitle", metavar="TEXT", help="Subtitle")
    meta.add_argument("--author", metavar="TEXT", help="Author")
    meta.add_argument("--date", metavar="TEXT", help="Date")
    meta.add_argument("--subject", metavar="TEXT", help="Subject")
    meta.add_argument("--keywords", metavar="TEXT", help="Keywords, comma-separated")

    advanced = parser.add_argument_group("Advanced")
    advanced.add_argument(
        "--reference-doc",
        metavar="FILE",
        help="Use an existing .docx as a style template (its content is discarded)",
    )

    return parser


# ----------------------------------------------------------------------
def _explicit_options(argv: Sequence[str], parser: argparse.ArgumentParser) -> set[str]:
    """Determines which Config fields were actually named on the command line."""
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
    """Builds the configuration from the parsed arguments."""
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

    # If the code background does not come from the theme, derive it from Pygments
    if "code_bg" not in values and "pygments_style" in explicit:
        from md2word.highlight import background_for_style

        derived = background_for_style(args.pygments_style)
        if derived:
            values["code_bg"] = derived

    config = Config(**values)
    config._explicit = explicit
    return config


def _expand_inputs(patterns: Sequence[str]) -> list[str]:
    """Expands glob patterns; the shell usually does this already."""
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
    if args.output:
        return args.output

    stem = "document" if input_path == "-" else os.path.splitext(os.path.basename(input_path))[0]
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
        parser.error("--output takes exactly one input file; use --output-dir instead")

    try:
        inputs = _expand_inputs(args.inputs)
    except FileNotFoundError as exc:
        print(f"Error: no file matches the pattern '{exc}'", file=sys.stderr)
        return 1

    if args.output and len(inputs) > 1:
        parser.error("--output takes exactly one input file; use --output-dir instead")

    explicit = _explicit_options(argv, parser)
    config = config_from_args(args, explicit)

    from md2word.converter import convert_file

    failures = 0
    for input_path in inputs:
        if input_path != "-" and not os.path.isfile(input_path):
            print(f"Error: file not found: {input_path}", file=sys.stderr)
            failures += 1
            continue

        output_path = _output_path(input_path, args)

        if os.path.exists(output_path) and not args.force:
            if os.path.abspath(output_path) == os.path.abspath(input_path):
                print(
                    f"Error: source and target are the same file: {output_path}",
                    file=sys.stderr,
                )
                failures += 1
                continue
            print(
                f"Skipped (already exists, use --force to overwrite): {output_path}",
                file=sys.stderr,
            )
            failures += 1
            continue

        try:
            result = convert_file(input_path, output_path, config)
        except FileNotFoundError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            failures += 1
            continue
        except Exception as exc:
            print(f"Error converting '{input_path}': {exc}", file=sys.stderr)
            if args.verbose:
                import traceback

                traceback.print_exc()
            failures += 1
            continue

        if not args.quiet:
            source = "standard input" if input_path == "-" else input_path
            size = os.path.getsize(result.output_path) / 1024
            headings = result.heading_count
            print(
                f"{source} -> {result.output_path}  "
                f"({size:.0f} KB, {headings} heading{'s' if headings != 1 else ''})"
            )
        for warning in result.warnings:
            print(f"  Note: {warning}", file=sys.stderr)

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
