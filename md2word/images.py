"""Bildquellen aufloesen: lokale Pfade, data:-URIs und entfernte URLs."""

from __future__ import annotations

import base64
import io
import os
import re
import urllib.parse
from dataclasses import dataclass

from docx.shared import Emu, Mm

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None  # type: ignore[assignment]

try:
    from PIL import Image as PILImage
except ImportError:  # pragma: no cover
    PILImage = None  # type: ignore[assignment]

from docx.image.image import Image as DocxImage

_DATA_URI = re.compile(r"^data:(?P<mime>[\w./+-]+)?(?:;charset=[\w-]+)?(?P<b64>;base64)?,(?P<data>.*)$", re.S)
_RASTER_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}

USER_AGENT = "md2word/1.0 (+https://pypi.org/project/python-docx/)"


class ImageError(RuntimeError):
    """Ein Bild konnte nicht geladen oder nicht eingebettet werden."""


@dataclass
class LoadedImage:
    stream: io.BytesIO
    width: Emu | None  # native Breite
    height: Emu | None
    source: str


def _is_remote(src: str) -> bool:
    return bool(re.match(r"^(https?|ftp)://", src, re.I))


def _decode_data_uri(src: str) -> bytes:
    match = _DATA_URI.match(src)
    if not match:
        raise ImageError("Ungueltige data:-URI")
    payload = match.group("data")
    if match.group("b64"):
        return base64.b64decode(payload)
    return urllib.parse.unquote_to_bytes(payload)


def _convert_svg(blob: bytes) -> bytes:
    """Wandelt SVG nach PNG, sofern cairosvg installiert ist."""
    try:
        import cairosvg  # type: ignore[import-not-found]
    except ImportError as exc:  # pragma: no cover
        raise ImageError(
            "SVG-Grafiken benoetigen das optionale Paket 'cairosvg' "
            "(pip install cairosvg)"
        ) from exc
    return cairosvg.svg2png(bytestring=blob, output_width=1600)


def _looks_like_svg(blob: bytes, source: str) -> bool:
    if source.lower().split("?")[0].endswith(".svg"):
        return True
    head = blob[:400].lstrip()
    return head.startswith(b"<svg") or (head.startswith(b"<?xml") and b"<svg" in blob[:2000])


def load_image(
    src: str,
    base_dir: str = ".",
    download: bool = True,
    timeout: float = 10.0,
) -> LoadedImage:
    """Laedt ein Bild und liefert Stream plus native Abmessungen."""
    src = (src or "").strip()
    if not src:
        raise ImageError("Leere Bildquelle")

    if src.startswith("data:"):
        blob = _decode_data_uri(src)
    elif _is_remote(src):
        if not download:
            raise ImageError(f"Entferntes Bild uebersprungen (--no-remote-images): {src}")
        if requests is None:
            raise ImageError("Fuer entfernte Bilder wird 'requests' benoetigt")
        response = requests.get(src, timeout=timeout, headers={"User-Agent": USER_AGENT})
        response.raise_for_status()
        blob = response.content
    else:
        path = src
        if src.startswith("file://"):
            path = urllib.parse.unquote(urllib.parse.urlparse(src).path)
        path = os.path.expanduser(path)
        if not os.path.isabs(path):
            path = os.path.join(base_dir, path)
        if not os.path.isfile(path):
            raise ImageError(f"Bilddatei nicht gefunden: {path}")
        with open(path, "rb") as handle:
            blob = handle.read()

    if _looks_like_svg(blob, src):
        blob = _convert_svg(blob)

    stream = io.BytesIO(blob)
    width = height = None
    try:
        info = DocxImage.from_blob(blob)
        width, height = info.width, info.height
    except Exception:  # python-docx kennt das Format nicht -> Pillow versuchen
        if PILImage is not None:
            try:
                with PILImage.open(io.BytesIO(blob)) as pil:
                    dpi = pil.info.get("dpi", (96, 96))
                    dpi_x = float(dpi[0] or 96)
                    dpi_y = float(dpi[1] or 96)
                    width = Emu(int(pil.width / dpi_x * 914400))
                    height = Emu(int(pil.height / dpi_y * 914400))
            except Exception as exc:
                raise ImageError(f"Bildformat nicht lesbar: {src} ({exc})") from exc
        else:
            raise ImageError(f"Bildformat nicht lesbar: {src}")

    stream.seek(0)
    return LoadedImage(stream=stream, width=width, height=height, source=src)


def fit_width(image: LoadedImage, max_width_mm: float) -> Emu | None:
    """Liefert die Zielbreite, wenn das Bild verkleinert werden muss."""
    limit = Mm(max_width_mm)
    if image.width is None:
        return limit
    return limit if image.width > limit else None
