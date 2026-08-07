#!/usr/bin/env python3
"""Baut md2word zu einem eigenstaendigen Programm (PyInstaller).

  python build.py              # Verzeichnisvariante, startet schnell
  python build.py --onefile    # eine einzelne Datei, bequem weiterzugeben
  python build.py --clean      # Zwischenergebnisse vorher loeschen

PyInstaller kann nicht fuer fremde Systeme bauen: die Windows-.exe muss
unter Windows entstehen, das macOS-Programm unter macOS, das Linux-Programm
unter Linux.
"""

from __future__ import annotations

import argparse
import os
import platform
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SPEC = ROOT / "md2word.spec"


def die(message: str) -> "NoReturn":  # type: ignore[valid-type]
    print(f"Fehler: {message}", file=sys.stderr)
    raise SystemExit(1)


def check_requirements() -> None:
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        die(
            "PyInstaller fehlt. Installation:\n"
            "  pip install pyinstaller\n"
            "oder gleich mit allen Bauwerkzeugen:\n"
            "  pip install -e \".[build]\""
        )

    missing = []
    for module, package in (
        ("docx", "python-docx"),
        ("markdown_it", "markdown-it-py"),
        ("mdit_py_plugins", "mdit-py-plugins"),
        ("lxml", "lxml"),
        ("pygments", "Pygments"),
        ("yaml", "PyYAML"),
    ):
        try:
            __import__(module)
        except ImportError:
            missing.append(package)
    if missing:
        die("Diese Pakete fehlen: " + ", ".join(missing) + "\n  pip install -e .")


def executable_name() -> str:
    return "md2word.exe" if platform.system() == "Windows" else "md2word"


def find_result(onefile: bool) -> Path | None:
    name = executable_name()
    candidate = ROOT / "dist" / name if onefile else ROOT / "dist" / "md2word" / name
    return candidate if candidate.exists() else None


def clear_previous_output() -> None:
    """Raeumt ein altes Ergebnis weg.

    Beim Wechsel zwischen Einzeldatei und Verzeichnis kollidieren beide
    Varianten unter demselben Namen in dist/ - PyInstaller bricht dann mit
    einem Rechtefehler ab.
    """
    for candidate in (ROOT / "dist" / "md2word", ROOT / "dist" / "md2word.exe"):
        if candidate.is_dir():
            shutil.rmtree(candidate, ignore_errors=True)
        elif candidate.exists():
            candidate.unlink()


def directory_size(path: Path) -> int:
    if path.is_file():
        return path.stat().st_size
    return sum(f.stat().st_size for f in path.rglob("*") if f.is_file())


def smoke_test(binary: Path) -> bool:
    """Prueft, ob das gebaute Programm wirklich konvertiert."""
    import tempfile

    sample = (
        "---\ntitle: Bauprobe\n---\n\n"
        "# Überschrift\n\nText mit **Auszeichnung**, `Code` und einer Fußnote[^1].\n\n"
        "- Liste\n- Zweiter Punkt\n\n"
        "| A | B |\n|---|---|\n| 1 | 2 |\n\n"
        "```python\ndef f():\n    return 42\n```\n\n"
        "[^1]: Die Anmerkung.\n"
    )

    with tempfile.TemporaryDirectory() as tmp:
        source = Path(tmp) / "probe.md"
        target = Path(tmp) / "probe.docx"
        source.write_text(sample, encoding="utf-8")

        result = subprocess.run(
            [str(binary), str(source), "-o", str(target), "--toc", "--page-numbers", "-q"],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            print("  Probelauf fehlgeschlagen:", result.stderr.strip()[:500])
            return False
        if not target.exists() or target.stat().st_size < 5000:
            print("  Probelauf lieferte keine brauchbare Datei")
            return False

        # Grobpruefung des Pakets
        import zipfile

        with zipfile.ZipFile(target) as archive:
            names = set(archive.namelist())
            for required in ("word/document.xml", "word/styles.xml", "word/numbering.xml"):
                if required not in names:
                    print(f"  Im Ergebnis fehlt {required}")
                    return False
            if "word/footnotes.xml" not in names:
                print("  Fussnoten fehlen im Ergebnis")
                return False
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--onefile",
        action="store_true",
        help="Alles in eine Datei packen (bequemer weiterzugeben, startet langsamer)",
    )
    parser.add_argument("--clean", action="store_true", help="build/ und dist/ vorher loeschen")
    parser.add_argument("--no-test", action="store_true", help="Probelauf nach dem Bauen ueberspringen")
    args = parser.parse_args()

    check_requirements()
    os.chdir(ROOT)

    if args.clean:
        for folder in ("build", "dist"):
            shutil.rmtree(ROOT / folder, ignore_errors=True)
        print("Zwischenergebnisse geloescht")

    clear_previous_output()

    command = [sys.executable, "-m", "PyInstaller", str(SPEC), "--noconfirm"]
    environment = dict(os.environ)
    if args.onefile:
        command += ["--", "--onefile"]
        environment["MD2WORD_ONEFILE"] = "1"
    else:
        environment.pop("MD2WORD_ONEFILE", None)

    print(f"Baue md2word fuer {platform.system()} {platform.machine()} "
          f"({'eine Datei' if args.onefile else 'ein Verzeichnis'}) ...")
    started = time.time()
    result = subprocess.run(command, env=environment)
    if result.returncode != 0:
        die("PyInstaller ist fehlgeschlagen")

    binary = find_result(args.onefile)
    if binary is None:
        die("Das gebaute Programm wurde nicht gefunden")

    size = directory_size(binary if args.onefile else binary.parent)
    print(f"\nFertig in {time.time() - started:.0f} s")
    print(f"  Programm : {binary}")
    print(f"  Groesse  : {size / 1024 / 1024:.1f} MB")

    if not args.no_test:
        print("\nProbelauf ...")
        if smoke_test(binary):
            print("  Probelauf bestanden - das Programm konvertiert korrekt.")
        else:
            die("Der Probelauf ist fehlgeschlagen")

    if platform.system() == "Darwin":
        print(
            "\nHinweis fuer macOS: Das Programm ist nicht signiert. Beim ersten\n"
            "Start kann Gatekeeper es blockieren. Freigeben mit:\n"
            f"  xattr -dr com.apple.quarantine {binary}"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
