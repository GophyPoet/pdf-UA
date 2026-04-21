"""Local OCR using Tesseract.

Returns not just the text but also a confidence figure, so downstream
code can honestly report when OCR is weak and tag it in the report.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageOps

try:
    import pytesseract
except Exception:  # pragma: no cover
    pytesseract = None  # type: ignore

log = logging.getLogger(__name__)


@dataclass
class OcrResult:
    text: str
    confidence: float  # 0..1
    lang: str
    word_count: int
    usable: bool  # confidence >= threshold

    def as_dict(self) -> dict:
        return dict(self.__dict__)


def _prepare(img: Image.Image) -> Image.Image:
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("L", "RGB"):
        img = img.convert("RGB")
    # Upscale small images — Tesseract is happier at ~300 DPI equivalents.
    w, h = img.size
    if max(w, h) < 1000:
        scale = max(1, int(1000 / max(w, h)))
        if scale > 1:
            img = img.resize((w * scale, h * scale), Image.LANCZOS)
    return img


def run_ocr(image_bytes: bytes, lang: str = "rus+eng", threshold: float = 0.55) -> OcrResult:
    if pytesseract is None:
        return OcrResult("", 0.0, lang, 0, False)
    try:
        img = Image.open(BytesIO(image_bytes))
    except Exception as e:
        log.warning("OCR cannot open image: %s", e)
        return OcrResult("", 0.0, lang, 0, False)
    img = _prepare(img)
    try:
        data = pytesseract.image_to_data(
            img, lang=lang, output_type=pytesseract.Output.DICT
        )
    except pytesseract.TesseractError as e:
        # Fall back to english-only if requested language is missing.
        if "rus" in lang and "eng" in lang:
            try:
                data = pytesseract.image_to_data(img, lang="eng", output_type=pytesseract.Output.DICT)
                lang = "eng"
            except Exception as e2:
                log.warning("OCR failed: %s", e2)
                return OcrResult("", 0.0, lang, 0, False)
        else:
            log.warning("OCR failed: %s", e)
            return OcrResult("", 0.0, lang, 0, False)
    words = []
    confs = []
    for i, w in enumerate(data.get("text", [])):
        if not w or not w.strip():
            continue
        try:
            c = float(data["conf"][i])
        except Exception:
            continue
        if c < 0:
            continue
        words.append(w.strip())
        confs.append(c / 100.0)
    text = " ".join(words).strip()
    # Keep line structure where possible by re-joining from image_to_string
    try:
        lined = pytesseract.image_to_string(img, lang=lang).strip()
        if lined:
            text = lined
    except Exception:
        pass
    conf = sum(confs) / len(confs) if confs else 0.0
    return OcrResult(
        text=text,
        confidence=round(conf, 3),
        lang=lang,
        word_count=len(words),
        usable=bool(words) and conf >= threshold,
    )
