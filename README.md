# md2word

A command-line tool that converts Markdown files into Microsoft Word documents
(`.docx`) — no Word installation, no Pandoc, pure Python.

The output is a genuine OOXML document with real styles, Word footnotes,
clickable cross-references, an updatable table of contents and page-number
fields. Word opens it without a repair prompt, and every bit of formatting stays
editable through the document's styles.

---

## Contents

- [What it converts](#what-it-converts)
- [Installation](#installation)
- [Usage](#usage)
- [Options](#options)
- [YAML front matter](#yaml-front-matter)
- [Building a standalone executable (PyInstaller)](#building-a-standalone-executable-pyinstaller)
- [Limitations](#limitations)
- [Development](#development)

---

## What it converts

| Markdown | Result in the Word document |
|:---------|:----------------------------|
| `#` through `######` | *Heading 1–6* styles, each with a bookmark for cross-references |
| `**bold**`, `*italic*`, `~~strikethrough~~` | Character formatting |
| `` `code` `` | *Verbatim Char* style with a monospace font and shading |
| ` ```python ` | Code block with syntax highlighting (Pygments), border and background |
| `- item` / `1. item` | Real Word lists, a fresh numbering sequence per list, up to 9 levels |
| `- [x] task` | Task list rendered with ☒/☐ |
| `> quote` | Quote style with a coloured bar, indented further when nested |
| `\| table \|` | Word table with a repeating header row, column alignment and fitted column widths |
| `[text](url)` | External hyperlink |
| `[text](#section)` | Internal link to the heading |
| `![image](file.png)` | Embedded image, scaled to the text width |
| `text[^1]` | A **real Word footnote** at the bottom of the page (not just superscript text) |
| `Term\n: Definition` | Definition list |
| `---` | Horizontal rule |
| `$a^2$`, `$$…$$` | **Real Word equations** (OMML), inline and as centred display blocks |
| `<b>HTML</b>` | Raw HTML is carried through (disable with `--strip-html`) |
| `<!-- pagebreak -->` | Hard page break |

On top of that: title page, table of contents, numbered headings, headers and
footers with page numbers, four themes, free choice of fonts and colours, paper
sizes from A5 to Legal, batch conversion, and document properties taken from the
front matter.

---

## Installation

Requires **Python 3.9 or newer**. Tested on 3.9.6 and 3.14.6 under macOS; the
code itself is platform-independent.

### Option A — as a Python package (recommended)

Cleanest inside a virtual environment, so the dependencies stay out of your
other projects:

```bash
python3 -m venv .venv
```

```bash
source .venv/bin/activate
```

On Windows the second line is `.venv\Scripts\activate`.

Then install the package:

```bash
pip install .
```

The `md2word` command is now available:

```bash
md2word --version
```

### Option B — straight from the source tree, without installing

Install only the dependencies and call the package as a module. You then have to
run it **from the project directory**, since `md2word` itself is not installed:

```bash
pip install -r requirements.txt
```

```bash
python -m md2word file.md
```

### Option C — system-wide with pipx

`pipx` installs the tool into its own environment while still putting the
command on your `PATH` — the most convenient route if you want md2word
available everywhere:

```bash
pipx install .
```

### Option D — a standalone executable, no Python at all

See [Building a standalone executable](#building-a-standalone-executable-pyinstaller).

### Optional extras

SVG graphics need one additional package, because Word cannot embed the format
and it has to be converted to PNG first:

```bash
pip install ".[svg]"
```

`cairosvg` depends on the Cairo system library — `brew install cairo` on macOS,
`apt install libcairo2` on Debian/Ubuntu. Without the extra, md2word skips SVG
images with a warning instead of failing. It is deliberately kept out of the
core dependencies so that a built executable does not depend on native libraries
of the target system.

---

## Usage

### The simplest case

```bash
md2word report.md
```

Writes `report.docx` next to the source file. An existing target file is **not**
overwritten — that takes `--force`.

### Choosing an output file or directory

```bash
md2word report.md -o ~/Desktop/Quarterly-Report.docx
```

```bash
md2word report.md --output-dir build
```

### Several files at once

```bash
md2word chapters/*.md --output-dir build --force
```

### Reading from a pipe

```bash
cat notes.md | md2word - -o notes.docx
```

### A fully equipped document

```bash
md2word manual.md --title-page --toc --number-headings --page-numbers --break-on-h1 --theme modern
```

That yields: a title page built from the metadata, a table of contents on its own
page, numbered headings (1., 1.1, 1.1.1), page numbers in the footer, and a page
break before every top-level chapter.

> **A note on the table of contents:** Word only evaluates field codes when the
> document is opened. md2word sets the flag that asks it to, so Word will offer
> to update the fields on first open — answer *Yes*. You can also do it any time
> later: right-click the table of contents → *Update Field* (or `Ctrl`/`Cmd` + `A`,
> then `F9`).

### Applying your own corporate design

The most elegant route is a reference document: an empty `.docx` in which the
*Normal*, *Heading 1–6*, *Quote* and similar styles already look the way you
want them.

```bash
md2word report.md --reference-doc template.docx
```

md2word adopts every style it finds there, along with the page layout, and only
adds what is missing (the code-block style, for instance). The template's
content is discarded. Because the styles carry the same names as in Pandoc's
reference document, a template built for Pandoc works here unchanged.

### Setting the appearance from the command line

```bash
md2word text.md --body-font "Georgia" --heading-font "Helvetica Neue" --font-size 11.5 --line-spacing 1.3 --accent 8C1D40
```

```bash
md2word text.md --page-size a5 --margin 15 --pygments-style dracula
```

List the available syntax colour schemes:

```bash
md2word --list-pygments-styles
```

---

## Options

### Input and output

| Option | Meaning |
|:-------|:--------|
| `FILE …` | One or more Markdown files; `-` reads from standard input |
| `-o`, `--output PATH` | Target file (only with exactly one input file) |
| `-d`, `--output-dir DIR` | Target directory for all output, created if needed |
| `-f`, `--force` | Overwrite existing target files |
| `-q`, `--quiet` | Suppress status messages |
| `-v`, `--verbose` | Verbose output including full tracebacks |
| `--version` | Print the version |

### Page layout

| Option | Default | Meaning |
|:-------|:--------|:--------|
| `--page-size {a3,a4,a5,letter,legal}` | `a4` | Paper size |
| `--landscape` | off | Landscape orientation |
| `--margin MM` | 25 | All four page margins |
| `--margin-top/-bottom/-left/-right MM` | 25 | A single margin; beats `--margin` |

### Typography

| Option | Default | Meaning |
|:-------|:--------|:--------|
| `--theme {default,classic,modern,mono}` | `default` | Colour and font scheme |
| `--body-font NAME` | from the theme | Body text font |
| `--heading-font NAME` | from the theme | Heading font |
| `--code-font NAME` | from the theme | Font for source code |
| `--font-size PT` | 11 | Base font size |
| `--code-font-size PT` | 9.5 | Font size inside code |
| `--line-spacing FACTOR` | 1.15 | Line spacing |
| `--accent HEX` | from the theme | Accent colour for rules and quote bars |
| `--link-color HEX` | from the theme | Hyperlink colour |
| `--code-bg HEX` | from the theme | Background of code blocks |

The four themes: `default` (Calibri, blue), `classic` (Times New Roman, black),
`modern` (Segoe UI, strong blue), `mono` (Arial, greyscale).

### Document structure

| Option | Meaning |
|:-------|:--------|
| `--toc` | Prepend a table of contents on its own page |
| `--toc-depth N` | Heading levels to include (default 3) |
| `--toc-title TEXT` | Heading above the table of contents (default follows `--lang`) |
| `--title-page` | Build a title page from title, subtitle, author and date |
| `--number-headings` | Number headings automatically (1., 1.1, 1.1.1) |
| `--page-numbers` | Page numbers as "page / total" in the footer |
| `--header-text TEXT` | Header text |
| `--footer-text TEXT` | Footer text |
| `--break-on-h1` | Page break before every level-1 heading |

### Content

| Option | Default | Meaning |
|:-------|:--------|:--------|
| `--no-highlight` | off | Turn off syntax highlighting |
| `--pygments-style NAME` | `friendly` | Colour scheme for source code |
| `--list-pygments-styles` | – | List the available colour schemes |
| `--no-remote-images` | off | Do not fetch images from `http(s)` addresses |
| `--max-image-width MM` | text width | Upper bound for image widths |
| `--captions {title,alt,none}` | `title` | What image captions are made from |
| `--footnotes {footnotes,endnotes}` | `footnotes` | Real footnotes, or notes collected at the end |
| `--math {omml,text}` | `omml` | Formulas as real Word equations, or as formatted text |
| `--strip-html` | off | Ignore raw HTML instead of carrying it through |
| `--lang CODE` | `en-US` | Document language; drives quotation marks and built-in headings |

On `--captions`: by default only the quoted title becomes a caption, so
`![alt](image.png "Figure 1")` → *Figure 1*. With `alt` the alternative text is
used as a fallback caption; with `none` there are no captions at all.

`--lang` does two things. It picks the quotation marks the typographer
inserts — `de` and related languages get „German“ ones, `fr` French « ones with
a narrow space », `es`, `it`, `ru` and others «guillemets without the space»,
everything else “English” ones. Non-breaking spaces in the source are preserved
either way.

It also selects the language of the handful of strings md2word writes *into* the
document: the table-of-contents heading, the endnotes heading, and the
placeholder shown for an image that could not be loaded. English, German,
French, Spanish and Italian are translated; any other language falls back to
English. Adding one is a single dictionary entry in `md2word/i18n.py`.

```bash
md2word bericht.md --lang de-DE    # „Inhaltsverzeichnis“, German quotes
```

### Metadata

`--title`, `--subtitle`, `--author`, `--date`, `--subject`, `--keywords`

These end up in the document properties and — with `--title-page` — on the title
page. Values given on the command line take precedence over the front matter.

### Advanced

`--reference-doc FILE` — use an existing `.docx` as a style template.

---

## YAML front matter

Metadata and most layout options can live in the file itself. That is handy
because the document then looks the same every time without a long command line:

```markdown
---
title: Quarterly Report Q3
subtitle: Status of the work packages
author: Jane Doe
date: 2026-08-07
keywords: project, report, Q3
lang: en-US

toc: true
toc_depth: 2
title_page: true
page_numbers: true
number_headings: true
theme: modern
page_size: a4
footer_text: Confidential — internal use only
---

# Background

The actual text starts here.
```

Recognised metadata: `title`, `subtitle`, `author` (a list works too), `date`,
`subject`/`description`, `keywords`/`tags`, `comments`, `lang`/`language`.
Recognised options: `toc`, `toc_depth`, `toc_title`, `title_page`,
`number_headings`, `page_numbers`, `header_text`, `footer_text`, `break_on_h1`,
`page_size`, `landscape`, `theme`, `font_size`, `body_font`, `heading_font`,
`code_font`, `highlight` and `pygments_style`.

Unknown keys are ignored silently, so your own project fields do nothing but sit
there. Anything given explicitly on the command line wins over the front matter.

---

## Building a standalone executable (PyInstaller)

Yes, this works — and it is tested. It lets md2word run on machines without a
Python installation.

### Building

```bash
pip install ".[build]"
```

```bash
python make_exe.py
```

The script checks the dependencies, runs PyInstaller against `md2word.spec`, and
then performs a smoke test: it converts a document containing headings, a list,
a table, a code block, a footnote and a table of contents, and verifies that the
result actually contains every expected part. Only then does the build count as
successful.

### The two build modes

```bash
python make_exe.py            # one directory (the default)
```

```bash
python make_exe.py --onefile  # a single file
```

Measured on an Apple Silicon Mac, each figure being one full conversion of the
bundled `examples/demo.md`:

| Mode | Result | Size | Time per invocation |
|:-----|:-------|-----:|--------------------:|
| Directory (default) | `dist/md2word/md2word` | 68 MB | **approx. 0.3 s** |
| Single file | `dist/md2word` | 21 MB | approx. 6.5 s |

The gap is not a measurement error: the single file unpacks itself into a
temporary folder on **every** start, and macOS inspects each of the roughly one
hundred bundled libraries as it does. On Windows and Linux the penalty is much
smaller, but still noticeable.

**Recommendation:** use the directory build for day-to-day work. The single file
pays off when you simply want to hand the program to someone — then convenience
outweighs startup time.

### Results per operating system

| System | File | How to use it |
|:-------|:-----|:--------------|
| Windows | `dist\md2word\md2word.exe` or `dist\md2word.exe` | Run it in a terminal, or add the folder to `PATH` |
| macOS | `dist/md2word/md2word` or `dist/md2word` | A Unix executable for the terminal |
| Linux | `dist/md2word/md2word` or `dist/md2word` | Same as macOS |

**PyInstaller cannot cross-build.** The Windows `.exe` has to be built on
Windows, the macOS executable on a Mac. To produce all three in one go you need
CI with matching runners (GitHub Actions with `windows-latest`, `macos-latest`
and `ubuntu-latest`, for example).

### Why no `.app` on macOS?

An `.app` bundle is meant for windowed programs. Double-clicking one starts it
without a terminal, so it can neither take arguments nor print anything — which
makes it useless for a command-line tool. The right form on macOS is the plain
Unix executable produced above. Treat it like any other command:

```bash
sudo cp -R dist/md2word /usr/local/lib/md2word && sudo ln -sf /usr/local/lib/md2word/md2word /usr/local/bin/md2word
```

After that `md2word` is available in any terminal.

### Gatekeeper on macOS

The built program is not signed with a Developer ID, so macOS may block it on
first launch. Clear the quarantine flag with:

```bash
xattr -dr com.apple.quarantine dist/md2word
```

If you intend to distribute the program, there is no way around signing and
notarising it with an Apple developer account (`codesign --deep --sign
"Developer ID Application: …"`, then `notarytool`).

### What matters in `md2word.spec`

Three pitfalls are already solved there — worth knowing if you adapt the spec:

1. **python-docx's base Word template** lives as a package file under
   `docx/templates/` and is picked up with `collect_data_files`. Without it,
   building any document fails.
2. **A placeholder in `docx/parts/`.** The modules there construct their
   template paths as `__file__ + "/../templates/…"`. In the bundle there is
   normally no file in `docx/parts/`, so the directory does not exist — and then
   the `..` cannot be resolved. Without the placeholder, the built program dies
   the moment a header, footer, page number or comment comes into play.
3. **Pygments** resolves lexers and colour schemes at runtime through name
   tables. They have to be registered as hidden imports via
   `collect_submodules`, otherwise there is no syntax highlighting at all.

---

## Limitations

- **Formulas** become real Word equations, editable in the equation editor. The
  translator covers the common ground — fractions, roots, sub- and
  superscripts, sums and integrals with limits, delimiters, accents, matrices,
  function names and Greek letters. Exotic LaTeX (custom macros, `\begin{align}`
  with alignment points, chemistry packages) is not understood; such a formula
  stays readable as formatted text and md2word says which one it was. Use
  `--math text` to switch equations off entirely.
- **SVG** cannot be embedded by Word directly. With the optional `cairosvg`
  package md2word converts SVG to PNG automatically; without it, the image is
  skipped.
- **Nested quotes** are rendered through increasing indentation, because Word
  has no real notion of nested block quotes.
- **Very wide tables** get squeezed into the text area. Beyond roughly ten
  columns, consider `--landscape` or a smaller `--font-size`.
- **Lists** support up to nine levels — that is Word's own limit. Deeper nesting
  collapses onto level nine.
- Word only refreshes the **table of contents and page numbers** when the
  document is opened or on request; that is by design, not a malfunction.

---

## Development

```bash
pip install -e ".[dev]"
```

```bash
pytest
```

The suite covers 252 tests across six files:

| File | Covers |
|:-----|:-------|
| `tests/test_blocks.py` | Headings, lists, quotes, code blocks, tables, rules |
| `tests/test_inline.py` | Character formatting, links, footnotes, images, raw HTML |
| `tests/test_document.py` | Front matter, title page, TOC, headers/footers, themes, reference documents |
| `tests/test_edge_cases.py` | Empty input, malformed markup, deep nesting, special characters |
| `tests/test_i18n.py` | Language fallbacks and the document strings that follow `--lang` |
| `tests/test_math.py` | LaTeX → OMML for every supported construct, plus the text fallback |

Every test that produces a complete file runs it through `assert_valid`. That
check, in `tests/conftest.py`, walks the resulting OPC package: well-formedness
of all XML parts, complete content types, resolvable relationships and `r:id`
references, existing styles, consistent numbering definitions including schema
order, paired bookmarks and field codes, footnotes that have a definition, and
anchors that actually point at a bookmark.

### Layout

| File | Responsibility |
|:-----|:---------------|
| `md2word/cli.py` | Command line, file names, batch processing |
| `md2word/config.py` | All settings and the themes |
| `md2word/parser.py` | Markdown → HTML (markdown-it-py), front matter |
| `md2word/renderer.py` | HTML tree → Word document |
| `md2word/converter.py` | Document frame: title page, TOC, headers and footers |
| `md2word/styles.py` | Styles and page layout |
| `md2word/docxutil.py` | Hand-written OOXML: footnotes, numbering, fields, borders |
| `md2word/images.py` | Resolving and embedding image sources |
| `md2word/highlight.py` | Syntax highlighting via Pygments |
| `md2word/i18n.py` | The few strings that go *into* the document, per language |
| `md2word/omml.py` | LaTeX → MathML → Word equations |

Routing through HTML is deliberate: nested structures — lists inside lists,
quotes containing code blocks, tables with inline markup — transfer far more
reliably from a tree than from the parser's flat token stream.

A full sample document exercising every supported element lives at
[examples/demo.md](examples/demo.md):

```bash
md2word examples/demo.md --toc --page-numbers --title-page --force
```

For the internals — how the pipeline works, why the OOXML is written by hand,
and where the traps are — see [SPEC.md](SPEC.md).

---

## License

MIT
