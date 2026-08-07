---
title: md2word – Feature Tour
subtitle: A sample document exercising every supported element
author: The md2word project
date: 2026-08-07
keywords: Markdown, Word, docx, conversion
lang: en-GB
---

# Introduction

This document shows what **md2word** makes of Markdown. It contains *italic*,
**bold**, ***bold italic*** and ~~struck-through~~ text, `inline code`, an
[external link](https://python-docx.readthedocs.io) and a reference to a
[later section](#tables).

A second paragraph with a footnote[^source] and a hard line break:  
this is what the second line looks like.

[^source]: Footnotes become **real** Word footnotes at the bottom of the page.

## Lists

- First item
- Second item with a nested list
  - Sub-item A
  - Sub-item B
    - One level deeper still
- Third item spanning several paragraphs

  This paragraph still belongs to the third item.

1. Numbered one
2. Numbered two
   1. Sub-item
   2. Another one
3. Numbered three

### Task list

- [x] Parser wired up
- [x] Styles defined
- [ ] Fetch coffee

## Quotes

> A plain quote running across several lines,
> which gets a coloured bar in the Word document.
>
> > And a nested quote inside it.

## Source code

A code block with syntax highlighting:

```python
from pathlib import Path


def count_words(path: Path) -> int:
    """Count the words in a text file."""
    text = path.read_text(encoding="utf-8")
    return len(text.split())


if __name__ == "__main__":
    print(count_words(Path("demo.md")))
```

And one without a language tag:

```
md2word demo.md --toc --page-numbers
```

## Tables

| Element      | Supported | Note                            |
|:-------------|:---------:|--------------------------------:|
| Headings     |    yes    | H1 through H6                   |
| Tables       |    yes    | including per-column alignment  |
| Footnotes    |    yes    | real Word footnotes             |
| Formulas     |    yes    | real Word equations (OMML)      |

## Definition lists

Markdown
: A lightweight markup language.

OOXML
: The XML format behind modern Office files.

## Rules and special characters

---

Accents and umlauts: äöü ÄÖÜ ß é à ç ñ ø å, quotation marks “English”,
„German“ and «French», an em dash – plus symbols: → ✓ € ½.

Non-breaking spaces survive conversion: 10&nbsp;kg, 25&nbsp;°C, page&nbsp;7.

Inline formula: $E = mc^2$

$$
\sum_{i=1}^{n} i = \frac{n(n+1)}{2}
$$

<!-- pagebreak -->

# Appendix

The text continues here after the page break. Raw HTML is carried through as
well: <b>bold via HTML</b>, <span style="color:#C00000">coloured</span>, and
<sup>superscript</sup>/<sub>subscript</sub>.
