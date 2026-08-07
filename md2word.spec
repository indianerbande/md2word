# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller-Bauplan fuer md2word.

Erzeugt ein eigenstaendiges Programm, das ohne Python-Installation laeuft:
  Windows -> dist/md2word.exe   bzw. dist/md2word/md2word.exe
  macOS   -> dist/md2word       (Unix-Executable fuer das Terminal)
  Linux   -> dist/md2word

Zwei Bauarten:
  pyinstaller md2word.spec --noconfirm                     # ein Verzeichnis (schneller Start)
  pyinstaller md2word.spec --noconfirm -- --onefile        # eine einzelne Datei

Bequemer geht es mit dem Hilfsskript:
  python build.py            # Verzeichnisvariante
  python build.py --onefile  # Einzeldatei
"""

import os
import sys

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# PyInstaller reicht alles nach "--" in sys.argv an die spec weiter (der
# Trenner selbst wird dabei entfernt). Umgebungsvariable als Ausweichweg.
ONEFILE = "--onefile" in sys.argv or os.environ.get("MD2WORD_ONEFILE") == "1"

# ----------------------------------------------------------------------
# Datendateien
# ----------------------------------------------------------------------
# python-docx bringt die Word-Grundvorlage als Paketdatei mit. Ohne sie
# scheitert jeder Document()-Aufruf im gebauten Programm.
datas = collect_data_files("docx", includes=["templates/*", "templates/**/*"])

# Die Module in docx/parts/ bauen ihre Vorlagenpfade als
# os.path.join(__file__, "..", "templates", ...). Im Bundle liegt in
# docx/parts/ keine Datei, also gibt es das Verzeichnis nicht - und das ".."
# laesst sich nicht aufloesen, wenn eine Pfadkomponente fehlt. Ein
# Platzhalter legt das Verzeichnis an und repariert damit Kopf- und
# Fusszeilen, Kommentare, Einstellungen und Formatvorlagen.
import docx as _docx  # noqa: E402  (nur zur Bauzeit importiert)

_marker = os.path.join(os.path.dirname(_docx.__file__), "py.typed")
if os.path.isfile(_marker):
    datas.append((_marker, "docx/parts"))

# ----------------------------------------------------------------------
# Versteckte Importe
# ----------------------------------------------------------------------
# Pygments laedt Lexer und Farbschemata erst zur Laufzeit ueber
# Namenstabellen - der statische Import-Scanner sieht sie nicht.
hiddenimports = (
    collect_submodules("pygments.lexers")
    + collect_submodules("pygments.styles")
    + [
        "mdit_py_plugins.front_matter",
        "mdit_py_plugins.footnote",
        "mdit_py_plugins.deflist",
        "mdit_py_plugins.tasklists",
        "mdit_py_plugins.dollarmath",
        "linkify_it",
        "uc_micro",
    ]
)

# Was md2word nicht braucht - haelt das Ergebnis klein
excludes = [
    "tkinter", "unittest", "pydoc", "doctest", "pdb",
    "numpy", "scipy", "pandas", "matplotlib", "IPython", "jupyter",
    "pytest", "setuptools", "pip", "wheel",
    "PyQt5", "PyQt6", "PySide2", "PySide6", "wx",
    "test", "distutils",
]

# ----------------------------------------------------------------------
a = Analysis(
    ["md2word/__main__.py"],
    pathex=[os.path.abspath(".")],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

if ONEFILE:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name="md2word",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=True,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=None,
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name="md2word",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=True,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=None,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name="md2word",
    )
