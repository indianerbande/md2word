#!/usr/bin/env python3
"""Builds md2word into a standalone executable (PyInstaller).

  python build.py              # directory build, starts fast
  python build.py --onefile    # a single file, convenient to hand around
  python build.py --clean      # remove intermediate output first

PyInstaller cannot cross-build: the Windows .exe has to be produced on
Windows, the macOS executable on macOS, the Linux one on Linux.
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
    print(f"Error: {message}", file=sys.stderr)
    raise SystemExit(1)


def check_requirements() -> None:
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        die(
            "PyInstaller is missing. Install it with:\n"
            "  pip install pyinstaller\n"
            "or pull in every build tool at once:\n"
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
        die("These packages are missing: " + ", ".join(missing) + "\n  pip install -e .")


def executable_name() -> str:
    return "md2word.exe" if platform.system() == "Windows" else "md2word"


def find_result(onefile: bool) -> Path | None:
    name = executable_name()
    candidate = ROOT / "dist" / name if onefile else ROOT / "dist" / "md2word" / name
    return candidate if candidate.exists() else None


def clear_previous_output() -> None:
    """Clears a previous result.

    Switching between single-file and directory builds makes both variants
    collide under the same name in dist/ - PyInstaller then aborts with a
    permission error.
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
    """Verifies that the built executable actually converts."""
    import tempfile

    sample = (
        "---\ntitle: Build check\n---\n\n"
        "# Heading\n\nText with **emphasis**, `code` and a footnote[^1].\n\n"
        "- List item\n- Second item\n\n"
        "| A | B |\n|---|---|\n| 1 | 2 |\n\n"
        "```python\ndef f():\n    return 42\n```\n\n"
        "[^1]: The note.\n"
    )

    with tempfile.TemporaryDirectory() as tmp:
        source = Path(tmp) / "check.md"
        target = Path(tmp) / "check.docx"
        source.write_text(sample, encoding="utf-8")

        result = subprocess.run(
            [str(binary), str(source), "-o", str(target), "--toc", "--page-numbers", "-q"],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            print("  smoke test failed:", result.stderr.strip()[:500])
            return False
        if not target.exists() or target.stat().st_size < 5000:
            print("  smoke test produced no usable file")
            return False

        # Rough check of the package
        import zipfile

        with zipfile.ZipFile(target) as archive:
            names = set(archive.namelist())
            for required in ("word/document.xml", "word/styles.xml", "word/numbering.xml"):
                if required not in names:
                    print(f"  the result is missing {required}")
                    return False
            if "word/footnotes.xml" not in names:
                print("  footnotes are missing from the result")
                return False
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--onefile",
        action="store_true",
        help="Pack everything into one file (easier to share, slower to start)",
    )
    parser.add_argument("--clean", action="store_true", help="Remove build/ and dist/ first")
    parser.add_argument("--no-test", action="store_true", help="Skip the smoke test after building")
    args = parser.parse_args()

    check_requirements()
    os.chdir(ROOT)

    if args.clean:
        for folder in ("build", "dist"):
            shutil.rmtree(ROOT / folder, ignore_errors=True)
        print("Removed intermediate output")

    clear_previous_output()

    command = [sys.executable, "-m", "PyInstaller", str(SPEC), "--noconfirm"]
    environment = dict(os.environ)
    if args.onefile:
        command += ["--", "--onefile"]
        environment["MD2WORD_ONEFILE"] = "1"
    else:
        environment.pop("MD2WORD_ONEFILE", None)

    print(f"Building md2word for {platform.system()} {platform.machine()} "
          f"({'single file' if args.onefile else 'directory'}) ...")
    started = time.time()
    result = subprocess.run(command, env=environment)
    if result.returncode != 0:
        die("PyInstaller failed")

    binary = find_result(args.onefile)
    if binary is None:
        die("could not find the built executable")

    size = directory_size(binary if args.onefile else binary.parent)
    print(f"\nDone in {time.time() - started:.0f} s")
    print(f"  Executable : {binary}")
    print(f"  Size       : {size / 1024 / 1024:.1f} MB")

    if not args.no_test:
        print("\nSmoke test ...")
        if smoke_test(binary):
            print("  smoke test passed - the executable converts correctly.")
        else:
            die("the smoke test failed")

    if platform.system() == "Darwin":
        print(
            "\nNote for macOS: the executable is not signed, so Gatekeeper may\n"
            "block it on first launch. Clear the quarantine flag with:\n"
            f"  xattr -dr com.apple.quarantine {binary}"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
