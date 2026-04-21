"""Alt text generation.

Rules:
  - Short, meaningful, useful for accessibility.
  - Not the generic "image" / "photo" alone.
  - If OCR found readable text → reflect that (e.g. "Скриншот документа. Первая строка: …").
  - If OCR hinted at a stamp / seal → reflect that.
  - If image is a small decorative bullet or line → mark decorative.
  - Otherwise a size-aware neutral description.

Alt text = short description. The full OCR text is kept in the report
and (when substantial) also injected into the document as a sibling
text block so the content isn't locked behind an image.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .ocr import OcrResult

STAMP_KEYWORDS = (
    "печать", "stamp", "подпись", "seal", "штамп",
)
SCREENSHOT_KEYWORDS = (
    "меню", "file", "edit", "кнопка", "окно", "window", "dialog",
)


@dataclass
class AltTextDecision:
    alt_text: str
    decorative: bool
    text_equivalent: str  # non-empty = inject as sibling block
    reasoning: str
    confidence: float  # 0..1

    def as_dict(self) -> dict:
        return dict(self.__dict__)


def _clean_first_line(text: str, limit: int = 80) -> str:
    line = ""
    for candidate in text.splitlines():
        candidate = candidate.strip()
        if len(candidate) >= 4:
            line = candidate
            break
    if not line:
        line = text.strip().split(". ")[0] if text.strip() else ""
    line = re.sub(r"\s+", " ", line)
    if len(line) > limit:
        line = line[: limit - 1].rstrip() + "…"
    return line


def decide(ocr: OcrResult, width_px: int, height_px: int) -> AltTextDecision:
    decorative = False
    confidence = ocr.confidence
    reasoning_parts = []

    is_tiny = max(width_px, height_px) < 80
    aspect = (width_px / max(1, height_px)) if height_px else 1
    is_rule = min(width_px, height_px) <= 6 and max(width_px, height_px) >= 40

    text = ocr.text.strip()
    lower = text.lower()

    if is_tiny or is_rule:
        return AltTextDecision(
            alt_text="",
            decorative=True,
            text_equivalent="",
            reasoning="tiny or rule-like image treated as decorative",
            confidence=0.9,
        )

    if ocr.usable and text:
        first_line = _clean_first_line(text)
        if any(k in lower for k in STAMP_KEYWORDS):
            alt = "Печать или штамп документа"
            if first_line:
                alt = f"Печать документа. Текст: {first_line}"
            eq = text if len(text) > 20 else ""
            reasoning_parts.append("OCR found stamp keywords")
        elif any(k in lower for k in SCREENSHOT_KEYWORDS) or len(text) > 120:
            alt = f"Скриншот документа или интерфейса. Первая строка: {first_line}"
            eq = text
            reasoning_parts.append("OCR found screenshot-like content; exporting text equivalent")
        else:
            alt = f"Изображение с текстом: {first_line}"
            eq = text if len(text) > 40 else ""
            reasoning_parts.append("OCR found readable text")
        return AltTextDecision(
            alt_text=alt,
            decorative=False,
            text_equivalent=eq,
            reasoning="; ".join(reasoning_parts),
            confidence=confidence,
        )

    if text and not ocr.usable:
        # Low-confidence OCR — expose uncertainty honestly.
        first_line = _clean_first_line(text)
        alt = f"Изображение, предположительно содержит текст: {first_line}"
        return AltTextDecision(
            alt_text=alt,
            decorative=False,
            text_equivalent="",  # do not inject unreliable text
            reasoning=f"low-confidence OCR ({confidence:.2f}); alt text marked uncertain",
            confidence=confidence,
        )

    # Nothing was extracted — still give a useful alt.
    if aspect > 2.5:
        alt = "Горизонтальная схема или диаграмма документа"
    elif aspect < 0.5:
        alt = "Вертикальная схема или иллюстрация документа"
    else:
        alt = "Иллюстрация документа"
    return AltTextDecision(
        alt_text=alt,
        decorative=False,
        text_equivalent="",
        reasoning="OCR returned nothing usable; generic size-aware alt text",
        confidence=0.3,
    )
