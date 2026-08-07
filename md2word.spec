# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build recipe for md2word.

Produces a standalone program that runs without a Python installation:
  Windows -> dist/md2word.exe   or dist/md2word/md2word.exe
  macOS   -> dist/md2word       (a Unix executable for the terminal)
  Linux   -> dist/md2word

Two build modes:
  pyinstaller md2word.spec --noconfirm                     # one directory (fast start)
  pyinstaller md2word.spec --noconfirm -- --onefile        # a single file

The helper script is more convenient:
  python make_exe.py            # directory build
  python make_exe.py --onefile  # single file
"""

import os
import sys

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# PyInstaller forwards everything after "--" in sys.argv to the spec (the
# separator itself is stripped). The env var is a fallback route.
ONEFILE = "--onefile" in sys.argv or os.environ.get("MD2WORD_ONEFILE") == "1"

# ----------------------------------------------------------------------
# Data files
# ----------------------------------------------------------------------
# python-docx ships the base Word template as a package file. Without it
# every Document() call in the built program fails.
datas = collect_data_files("docx", includes=["templates/*", "templates/**/*"])

# The modules in docx/parts/ build their template paths as
# os.path.join(__file__, "..", "templates", ...). The bundle has no file in
# docx/parts/, so the directory does not exist - and ".." cannot be
# resolved when a path component is missing. A placeholder creates the
# directory and thereby fixes headers, footers, comments, settings and
# styles.
import docx as _docx  # noqa: E402  (imported at build time only)

_marker = os.path.join(os.path.dirname(_docx.__file__), "py.typed")
if os.path.isfile(_marker):
    datas.append((_marker, "docx/parts"))

# latex2mathml looks up every symbol in unimathsymbols.txt at runtime. It is a
# data file, so the import scanner does not see it, and without it no formula
# converts in the built program.
datas += collect_data_files("latex2mathml")

# ----------------------------------------------------------------------
# Hidden imports
# ----------------------------------------------------------------------
# Pygments resolves lexers and colour schemes at runtime through name
# tables - the static import scanner never sees them.
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

# What md2word does not need - keeps the result small
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
