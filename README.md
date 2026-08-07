# md2word

Kommandozeilen-Tool, das Markdown-Dateien in Microsoft-Word-Dokumente (`.docx`)
konvertiert – ohne installiertes Word, ohne Pandoc, rein in Python.

Das Ergebnis ist ein echtes OOXML-Dokument mit Formatvorlagen, Word-Fußnoten,
klickbaren Querverweisen, aktualisierbarem Inhaltsverzeichnis und
Seitenzahlfeldern. Word öffnet es ohne Reparaturhinweis, und alle Formatierungen
bleiben nachträglich über die Formatvorlagen änderbar.

---

## Inhalt

- [Funktionsumfang](#funktionsumfang)
- [Installation](#installation)
- [Aufruf](#aufruf)
- [Optionen](#optionen)
- [YAML-Front-Matter](#yaml-front-matter)
- [Eigenständiges Programm bauen (PyInstaller)](#eigenständiges-programm-bauen-pyinstaller)
- [Grenzen](#grenzen)
- [Entwicklung](#entwicklung)

---

## Funktionsumfang

| Markdown | Ergebnis im Word-Dokument |
|:---------|:--------------------------|
| `#` bis `######` | Formatvorlagen *Überschrift 1–6*, mit Lesezeichen für Querverweise |
| `**fett**`, `*kursiv*`, `~~durchgestrichen~~` | Zeichenformatierung |
| `` `Code` `` | Zeichenformat *Verbatim Char* mit Monospace-Schrift und Hintergrund |
| ` ```python ` | Codeblock mit Syntaxhervorhebung (Pygments), Rahmen und Hintergrund |
| `- Punkt` / `1. Punkt` | Echte Word-Listen mit eigener Nummerierung je Liste, bis 9 Ebenen |
| `- [x] Aufgabe` | Aufgabenliste mit ☒/☐ |
| `> Zitat` | Zitatformat mit farbigem Balken, gestaffelt bei Verschachtelung |
| `\| Tabelle \|` | Word-Tabelle mit Kopfzeilenwiederholung, Spaltenausrichtung und passenden Spaltenbreiten |
| `[Text](url)` | Externer Hyperlink |
| `[Text](#abschnitt)` | Interner Verweis auf die Überschrift |
| `![Bild](datei.png)` | Eingebettetes Bild, auf die Textbreite skaliert |
| `Text[^1]` | **Echte Word-Fußnote** am Seitenende (nicht nur hochgestellter Text) |
| `Begriff\n: Erklärung` | Definitionsliste |
| `---` | Trennlinie |
| `$a^2$`, `$$…$$` | Formeln als kursiver Cambria-Math-Text (siehe [Grenzen](#grenzen)) |
| `<b>HTML</b>` | Rohes HTML wird übernommen (abschaltbar mit `--strip-html`) |
| `<!-- pagebreak -->` | Harter Seitenumbruch |

Dazu kommen: Titelseite, Inhaltsverzeichnis, nummerierte Überschriften,
Kopf-/Fußzeilen mit Seitenzahlen, vier Farbschemata, freie Schrift- und
Farbwahl, Papierformate von A5 bis Legal, Stapelverarbeitung und
Dokumenteigenschaften aus dem Front Matter.

---

## Installation

Vorausgesetzt wird **Python 3.9 oder neuer**. Getestet mit 3.9.6 und 3.14.6 auf
macOS; der Code ist plattformunabhängig.

### Variante A – als Python-Paket (empfohlen)

Am saubersten in einer virtuellen Umgebung, damit die Abhängigkeiten nicht mit
anderen Projekten kollidieren:

```bash
python3 -m venv .venv
```

```bash
source .venv/bin/activate
```

Unter Windows lautet die zweite Zeile `.venv\Scripts\activate`.

Dann das Paket installieren:

```bash
pip install .
```

Danach steht der Befehl `md2word` zur Verfügung:

```bash
md2word --version
```

### Variante B – ohne Installation direkt aus dem Quellverzeichnis

Nur die Abhängigkeiten installieren und das Paket als Modul aufrufen. Der Aufruf
muss dann **aus dem Projektverzeichnis** heraus erfolgen, weil `md2word` selbst
nicht installiert wird:

```bash
pip install -r requirements.txt
```

```bash
python -m md2word datei.md
```

### Variante C – systemweit mit pipx

`pipx` installiert das Tool in eine eigene Umgebung und legt den Befehl trotzdem
global ab – der bequemste Weg, wenn md2word überall verfügbar sein soll:

```bash
pipx install .
```

### Variante D – eigenständiges Programm, ganz ohne Python

Siehe [Eigenständiges Programm bauen](#eigenständiges-programm-bauen-pyinstaller).

### Optionale Zusätze

SVG-Grafiken benötigen ein zusätzliches Paket, weil Word das Format nicht
einbetten kann und es vorher in PNG umgewandelt werden muss:

```bash
pip install ".[svg]"
```

`cairosvg` setzt die Systembibliothek Cairo voraus – unter macOS
`brew install cairo`, unter Debian/Ubuntu `apt install libcairo2`. Ohne das
Extra überspringt md2word SVG-Bilder mit einem Hinweis, statt abzubrechen.
Bewusst nicht in den Grundabhängigkeiten, damit ein gebautes Programm nicht von
nativen Bibliotheken des Zielsystems abhängt.

---

## Aufruf

### Der einfachste Fall

```bash
md2word bericht.md
```

Erzeugt `bericht.docx` neben der Quelldatei. Eine bereits vorhandene Zieldatei
wird **nicht** überschrieben – dafür braucht es `--force`.

### Zieldatei oder Zielordner angeben

```bash
md2word bericht.md -o ~/Desktop/Quartalsbericht.docx
```

```bash
md2word bericht.md --output-dir build
```

### Mehrere Dateien auf einmal

```bash
md2word kapitel/*.md --output-dir build --force
```

### Aus einer Pipe lesen

```bash
cat notizen.md | md2word - -o notizen.docx
```

### Ein vollständig ausgestattetes Dokument

```bash
md2word handbuch.md --title-page --toc --number-headings --page-numbers --break-on-h1 --theme modern
```

Das ergibt: Titelseite aus den Metadaten, Inhaltsverzeichnis auf eigener Seite,
durchnummerierte Überschriften (1., 1.1, 1.1.1), Seitenzahlen in der Fußzeile
und einen Seitenumbruch vor jedem Hauptkapitel.

> **Hinweis zum Inhaltsverzeichnis:** Word berechnet Feldfunktionen erst beim
> Öffnen. md2word setzt das Kennzeichen dafür, sodass Word beim ersten Öffnen
> fragt, ob die Felder aktualisiert werden sollen – hier mit *Ja* antworten.
> Nachträglich geht es jederzeit mit Rechtsklick auf das Verzeichnis →
> *Felder aktualisieren* (oder `Strg`/`Cmd` + `A`, dann `F9`).

### Ein eigenes Corporate Design verwenden

Am elegantesten über ein Referenzdokument: eine leere `.docx`, in der die
Formatvorlagen *Standard*, *Überschrift 1–6*, *Zitat* usw. bereits so aussehen,
wie sie sollen.

```bash
md2word bericht.md --reference-doc vorlage.docx
```

md2word übernimmt daraus alle vorhandenen Formatvorlagen samt Seitenlayout
unverändert und ergänzt nur, was noch fehlt (etwa das Codeblock-Format). Der
Inhalt der Vorlage wird verworfen. Da die Formatvorlagen dieselben Namen tragen
wie in Pandocs Referenzdokument, lässt sich ein für Pandoc erstelltes Template
direkt weiterverwenden.

### Aussehen direkt auf der Kommandozeile bestimmen

```bash
md2word text.md --body-font "Georgia" --heading-font "Helvetica Neue" --font-size 11.5 --line-spacing 1.3 --accent 8C1D40
```

```bash
md2word text.md --page-size a5 --margin 15 --pygments-style dracula
```

Verfügbare Farbschemata für Quelltext auflisten:

```bash
md2word --list-pygments-styles
```

---

## Optionen

### Ein- und Ausgabe

| Option | Bedeutung |
|:-------|:----------|
| `DATEI …` | Eine oder mehrere Markdown-Dateien; `-` liest von der Standardeingabe |
| `-o`, `--output PFAD` | Zieldatei (nur bei genau einer Eingabedatei) |
| `-d`, `--output-dir ORDNER` | Zielordner für alle Ausgaben, wird bei Bedarf angelegt |
| `-f`, `--force` | Vorhandene Zieldateien überschreiben |
| `-q`, `--quiet` | Keine Statusmeldungen |
| `-v`, `--verbose` | Ausführliche Ausgabe samt vollständigem Fehlerbericht |
| `--version` | Version ausgeben |

### Seitenlayout

| Option | Standard | Bedeutung |
|:-------|:---------|:----------|
| `--page-size {a3,a4,a5,letter,legal}` | `a4` | Papierformat |
| `--landscape` | aus | Querformat |
| `--margin MM` | 25 | Alle vier Seitenränder |
| `--margin-top/-bottom/-left/-right MM` | 25 | Einzelner Rand, schlägt `--margin` |

### Typografie

| Option | Standard | Bedeutung |
|:-------|:---------|:----------|
| `--theme {default,classic,modern,mono}` | `default` | Farb- und Schriftschema |
| `--body-font NAME` | aus dem Schema | Schriftart des Fließtextes |
| `--heading-font NAME` | aus dem Schema | Schriftart der Überschriften |
| `--code-font NAME` | aus dem Schema | Schriftart für Quelltext |
| `--font-size PT` | 11 | Grundschriftgröße |
| `--code-font-size PT` | 9.5 | Schriftgröße im Code |
| `--line-spacing FAKTOR` | 1.15 | Zeilenabstand |
| `--accent HEX` | aus dem Schema | Akzentfarbe für Linien und Zitatbalken |
| `--link-color HEX` | aus dem Schema | Farbe für Hyperlinks |
| `--code-bg HEX` | aus dem Schema | Hintergrund von Codeblöcken |

Die vier Schemata: `default` (Calibri, Blau), `classic` (Times New Roman,
Schwarz), `modern` (Segoe UI, kräftiges Blau), `mono` (Arial, Graustufen).

### Dokumentstruktur

| Option | Bedeutung |
|:-------|:----------|
| `--toc` | Inhaltsverzeichnis auf eigener Seite voranstellen |
| `--toc-depth N` | Erfasste Überschriftenebenen (Standard 3) |
| `--toc-title TEXT` | Überschrift des Verzeichnisses (Standard „Inhaltsverzeichnis") |
| `--title-page` | Titelseite aus Titel, Untertitel, Verfasser und Datum |
| `--number-headings` | Überschriften automatisch nummerieren (1., 1.1, 1.1.1) |
| `--page-numbers` | Seitenzahlen als „Seite / Gesamt" in die Fußzeile |
| `--header-text TEXT` | Text der Kopfzeile |
| `--footer-text TEXT` | Text der Fußzeile |
| `--break-on-h1` | Seitenumbruch vor jeder Überschrift der Ebene 1 |

### Inhalte

| Option | Standard | Bedeutung |
|:-------|:---------|:----------|
| `--no-highlight` | aus | Syntaxhervorhebung abschalten |
| `--pygments-style NAME` | `friendly` | Farbschema für Quelltext |
| `--list-pygments-styles` | – | Verfügbare Farbschemata auflisten |
| `--no-remote-images` | aus | Bilder von `http(s)`-Adressen nicht laden |
| `--max-image-width MM` | Textbreite | Obergrenze für Bildbreiten |
| `--captions {title,alt,none}` | `title` | Woraus Bildunterschriften entstehen |
| `--footnotes {footnotes,endnotes}` | `footnotes` | Echte Fußnoten oder gesammelte Anmerkungen am Ende |
| `--strip-html` | aus | Rohes HTML ignorieren statt übernehmen |
| `--lang CODE` | `de-DE` | Dokumentsprache; steuert auch die Anführungszeichen |

Zu `--captions`: Standardmäßig wird nur der Titel in Anführungszeichen zur
Bildunterschrift, also `![Alt](bild.png "Abbildung 1")` → *Abbildung 1*. Mit
`alt` dient ersatzweise der Alternativtext als Unterschrift, mit `none` gibt es
gar keine.

Zur Sprache: `de` und verwandte Sprachen bekommen „deutsche“ Anführungszeichen,
`fr` französische « mit schmalem Abstand », `es`, `it`, `ru` und weitere
«Guillemets ohne Abstand», alle übrigen “englische”. Geschützte Leerzeichen aus
der Quelle bleiben dabei erhalten.

### Metadaten

`--title`, `--subtitle`, `--author`, `--date`, `--subject`, `--keywords`

Landen in den Dokumenteigenschaften und – bei `--title-page` – auf der
Titelseite. Angaben auf der Kommandozeile haben Vorrang vor dem Front Matter.

### Erweitert

`--reference-doc DATEI` – bestehendes `.docx` als Formatvorlage.

---

## YAML-Front-Matter

Metadaten und die meisten Layoutoptionen lassen sich auch in die Datei selbst
schreiben. Das ist praktisch, weil das Dokument dann ohne lange Befehlszeile
immer gleich aussieht:

```markdown
---
title: Projektbericht Q3
subtitle: Stand der Arbeitspakete
author: Jane Doe
date: 2026-08-07
keywords: Projekt, Bericht, Q3
lang: de-AT

toc: true
toc_depth: 2
title_page: true
page_numbers: true
number_headings: true
theme: modern
page_size: a4
footer_text: Vertraulich – nur für den internen Gebrauch
---

# Ausgangslage

Der eigentliche Text beginnt hier.
```

Erkannt werden die Metadaten `title`, `subtitle`, `author` (auch als Liste),
`date`, `subject`/`description`, `keywords`/`tags`, `comments`, `lang`/`language`
sowie die Optionen `toc`, `toc_depth`, `toc_title`, `title_page`,
`number_headings`, `page_numbers`, `header_text`, `footer_text`, `break_on_h1`,
`page_size`, `landscape`, `theme`, `font_size`, `body_font`, `heading_font`,
`code_font`, `highlight` und `pygments_style`.

Unbekannte Schlüssel werden stillschweigend ignoriert – eigene Projektfelder
stören also nicht. Was auf der Kommandozeile explizit angegeben wird, gewinnt
gegen das Front Matter.

---

## Eigenständiges Programm bauen (PyInstaller)

Ja, das funktioniert – und ist getestet. Damit läuft md2word auf Rechnern ohne
Python-Installation.

### Bauen

```bash
pip install ".[build]"
```

```bash
python build.py
```

Das Skript prüft die Abhängigkeiten, ruft PyInstaller mit `md2word.spec` auf und
macht anschließend einen Probelauf: Es konvertiert ein Testdokument mit
Überschriften, Liste, Tabelle, Codeblock, Fußnote und Inhaltsverzeichnis und
prüft, ob im Ergebnis alle erwarteten Bestandteile stecken. Erst dann gilt der
Bau als erfolgreich.

### Die zwei Bauarten

```bash
python build.py            # ein Verzeichnis (Voreinstellung)
```

```bash
python build.py --onefile  # eine einzelne Datei
```

Gemessen auf einem Mac mit Apple Silicon, jeweils eine vollständige
Konvertierung des mitgelieferten `examples/demo.md`:

| Bauart | Ergebnis | Größe | Zeit pro Aufruf |
|:-------|:---------|------:|----------------:|
| Verzeichnis (Standard) | `dist/md2word/md2word` | 68 MB | **ca. 0,3 s** |
| Einzeldatei | `dist/md2word` | 21 MB | ca. 6,5 s |

Der große Unterschied ist kein Messfehler: Die Einzeldatei entpackt sich bei
**jedem** Start in einen temporären Ordner, und macOS prüft dabei jede der rund
hundert enthaltenen Bibliotheken einzeln. Auf Windows und Linux fällt der
Aufschlag deutlich geringer aus, bleibt aber spürbar.

**Empfehlung:** Für die tägliche Arbeit die Verzeichnisvariante. Die Einzeldatei
lohnt sich, wenn das Programm einfach weitergegeben werden soll – dann zählt
Bequemlichkeit mehr als die Startzeit.

### Ergebnisse je Betriebssystem

| System | Datei | Verwendung |
|:-------|:------|:-----------|
| Windows | `dist\md2word\md2word.exe` bzw. `dist\md2word.exe` | Direkt im Terminal, oder Ordner in `PATH` aufnehmen |
| macOS | `dist/md2word/md2word` bzw. `dist/md2word` | Unix-Programm für das Terminal |
| Linux | `dist/md2word/md2word` bzw. `dist/md2word` | Wie unter macOS |

**PyInstaller kann nicht für fremde Systeme bauen.** Die Windows-`.exe` muss auf
einem Windows-Rechner entstehen, das macOS-Programm auf einem Mac. Für alle drei
Plattformen aus einem Rutsch braucht es eine CI mit entsprechenden Runnern
(z. B. GitHub Actions mit `windows-latest`, `macos-latest` und `ubuntu-latest`).

### Warum unter macOS keine `.app`?

Ein `.app`-Bundle ist für Programme mit Fenster gedacht. Beim Doppelklick
startet es ohne Terminal, kann also weder Argumente entgegennehmen noch etwas
ausgeben – für ein Kommandozeilenwerkzeug ist das nutzlos. Die richtige Form
unter macOS ist das schlichte Unix-Programm, das oben entsteht. Es lässt sich
wie jedes andere Kommando behandeln:

```bash
sudo cp -R dist/md2word /usr/local/lib/md2word && sudo ln -sf /usr/local/lib/md2word/md2word /usr/local/bin/md2word
```

Danach ist `md2word` überall im Terminal verfügbar.

### Gatekeeper unter macOS

Das gebaute Programm ist nicht mit einer Entwickler-ID signiert. Beim ersten
Start blockiert macOS es deshalb möglicherweise. Freigeben mit:

```bash
xattr -dr com.apple.quarantine dist/md2word
```

Soll das Programm weitergegeben werden, führt an einer Signatur und
Notarisierung mit einem Apple-Entwicklerkonto kein Weg vorbei
(`codesign --deep --sign "Developer ID Application: …"`, danach `notarytool`).

### Was in der `md2word.spec` wichtig ist

Zwei Stolperstellen sind dort bereits gelöst – wer die Spec anpasst, sollte sie
kennen:

1. **Die Word-Grundvorlage von python-docx** liegt als Paketdatei unter
   `docx/templates/` und wird mit `collect_data_files` eingesammelt. Ohne sie
   scheitert jeder Dokumentaufbau.
2. **Ein Platzhalter in `docx/parts/`.** Die Module dort bilden ihre
   Vorlagenpfade als `__file__ + "/../templates/…"`. Im Bundle liegt in
   `docx/parts/` normalerweise keine Datei, das Verzeichnis existiert also
   nicht – und dann lässt sich das `..` nicht auflösen. Ohne diesen Platzhalter
   bricht das gebaute Programm ab, sobald eine Kopf- oder Fußzeile,
   Seitenzahlen oder Kommentare ins Spiel kommen.
3. **Pygments** lädt Lexer und Farbschemata erst zur Laufzeit über
   Namenstabellen. Sie müssen mit `collect_submodules` als versteckte Importe
   angemeldet werden, sonst gibt es keine Syntaxhervorhebung.

---

## Grenzen

- **Formeln** werden als kursiver Text in Cambria Math übernommen, nicht als
  bearbeitbare Word-Formel (OMML). Für gesetzte Formeln ist Pandoc mit
  `--to docx` die bessere Wahl. md2word weist beim Konvertieren darauf hin.
- **SVG** kann Word nicht direkt einbetten. Mit dem optionalen Paket `cairosvg`
  wandelt md2word SVG automatisch in PNG um, ohne bleibt das Bild aus.
- **Verschachtelte Zitate** werden über zunehmende Einzüge dargestellt, weil
  Word keine echte Zitatverschachtelung kennt.
- **Sehr breite Tabellen** werden auf den Satzspiegel gestaucht. Ab etwa zehn
  Spalten empfiehlt sich `--landscape` oder ein kleinerer `--font-size`.
- **Listen** unterstützen bis zu neun Ebenen – das ist die Grenze von Word
  selbst. Tiefere Verschachtelungen landen auf Ebene neun.
- Word aktualisiert **Inhaltsverzeichnis und Seitenzahlen** erst beim Öffnen
  bzw. auf Anforderung; das ist so vorgesehen und keine Fehlfunktion.

---

## Entwicklung

```bash
pip install -e ".[dev]"
```

```bash
pytest
```

Die Testsuite umfasst 193 Tests in vier Dateien:

| Datei | Prüft |
|:------|:------|
| `tests/test_blocks.py` | Überschriften, Listen, Zitate, Codeblöcke, Tabellen, Trennlinien |
| `tests/test_inline.py` | Zeichenformate, Links, Fußnoten, Bilder, rohes HTML |
| `tests/test_document.py` | Front Matter, Titelseite, Verzeichnis, Kopf-/Fußzeile, Schemata, Referenzdokument |
| `tests/test_edge_cases.py` | Leere Eingaben, kaputtes Markup, tiefe Verschachtelung, Sonderzeichen |

Jeder Test, der eine ganze Datei erzeugt, lässt sie durch `assert_valid` laufen.
Diese Prüfung in `tests/conftest.py` geht das erzeugte OPC-Paket durch:
Wohlgeformtheit aller XML-Teile, vollständige Content-Types, auflösbare
Beziehungen und `r:id`-Verweise, vorhandene Formatvorlagen, konsistente
Nummerierungen samt Schema-Reihenfolge, paarige Lesezeichen und Feldfunktionen,
Fußnoten mit Definition und Anker, auf die tatsächlich ein Lesezeichen zeigt.

### Aufbau

| Datei | Aufgabe |
|:------|:--------|
| `md2word/cli.py` | Kommandozeile, Dateinamen, Stapelverarbeitung |
| `md2word/config.py` | Alle Einstellungen und die Farbschemata |
| `md2word/parser.py` | Markdown → HTML (markdown-it-py), Front Matter |
| `md2word/renderer.py` | HTML-Baum → Word-Dokument |
| `md2word/converter.py` | Dokumentrahmen: Titelseite, Verzeichnis, Kopf-/Fußzeile |
| `md2word/styles.py` | Formatvorlagen und Seitenlayout |
| `md2word/docxutil.py` | OOXML-Handarbeit: Fußnoten, Nummerierung, Felder, Rahmen |
| `md2word/images.py` | Bildquellen auflösen und einbetten |
| `md2word/highlight.py` | Syntaxhervorhebung via Pygments |

Der Weg über HTML ist Absicht: Verschachtelte Strukturen – Listen in Listen,
Zitate mit Codeblöcken, Tabellen mit Auszeichnung – lassen sich aus einem Baum
wesentlich zuverlässiger übertragen als aus dem flachen Token-Strom des Parsers.

Ein vollständiges Beispieldokument mit allen unterstützten Elementen liegt unter
[examples/demo.md](examples/demo.md):

```bash
md2word examples/demo.md --toc --page-numbers --title-page --force
```

---

## Lizenz

MIT
