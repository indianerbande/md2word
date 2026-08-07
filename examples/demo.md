---
title: md2word – Funktionsübersicht
subtitle: Ein Beispieldokument mit allen unterstützten Elementen
author: Jürgen Reitböck
date: 2026-08-07
keywords: Markdown, Word, docx, Konvertierung
lang: de-AT
---

# Einleitung

Dieses Dokument zeigt, was **md2word** aus Markdown macht. Es enthält *kursiven*,
**fetten**, ***fett-kursiven*** und ~~durchgestrichenen~~ Text, `Inline-Code`
sowie einen [externen Link](https://python-docx.readthedocs.io) und einen
Verweis auf einen [späteren Abschnitt](#tabellen).

Ein zweiter Absatz mit einer Fußnote[^quelle] und einem harten
Zeilenumbruch:  
So sieht die zweite Zeile aus.

[^quelle]: Fußnoten landen als **echte** Word-Fußnoten am Seitenende.

## Aufzählungen

- Erster Punkt
- Zweiter Punkt mit verschachtelter Liste
  - Untereintrag A
  - Untereintrag B
    - Noch eine Ebene tiefer
- Dritter Punkt mit mehreren Absätzen

  Dieser Absatz gehört noch zum dritten Punkt.

1. Nummeriert eins
2. Nummeriert zwei
   1. Untereintrag
   2. Noch einer
3. Nummeriert drei

### Aufgabenliste

- [x] Parser angebunden
- [x] Formatvorlagen definiert
- [ ] Kaffee holen

## Zitate

> Ein einfaches Zitat über mehrere Zeilen hinweg,
> das im Word-Dokument einen farbigen Balken bekommt.
>
> > Und ein verschachteltes Zitat darin.

## Quelltext

Ein Codeblock mit Syntaxhervorhebung:

```python
from pathlib import Path


def zaehle_woerter(pfad: Path) -> int:
    """Zählt die Wörter einer Textdatei."""
    text = pfad.read_text(encoding="utf-8")
    return len(text.split())


if __name__ == "__main__":
    print(zaehle_woerter(Path("demo.md")))
```

Und einer ohne Sprachangabe:

```
$ md2word demo.md --toc --page-numbers
demo.md -> demo.docx  (28 KB, 12 Überschriften)
```

## Tabellen

| Element        | Unterstützt | Anmerkung                          |
|:---------------|:-----------:|-----------------------------------:|
| Überschriften  |     ja      | H1 bis H6                          |
| Tabellen       |     ja      | inkl. Ausrichtung pro Spalte       |
| Fußnoten       |     ja      | echte Word-Fußnoten                |
| Formeln        |  teilweise  | als formatierter Text              |

## Definitionslisten

Markdown
: Eine leichtgewichtige Auszeichnungssprache.

OOXML
: Das XML-Format hinter modernen Office-Dateien.

## Trennlinie und Sonderzeichen

---

Umlaute äöü ÄÖÜ ß, Anführungszeichen „deutsch“ und »französisch«,
Gedankenstrich – sowie Symbole: → ✓ € ½.

Inline-Formel: $E = mc^2$

$$
\sum_{i=1}^{n} i = \frac{n(n+1)}{2}
$$

<!-- pagebreak -->

# Anhang

Nach dem Seitenumbruch geht es hier weiter. Rohes HTML wird ebenfalls
übernommen: <b>fett per HTML</b>, <span style="color:#C00000">farbig</span>
und <sup>hochgestellt</sup>/<sub>tiefgestellt</sub>.
