"""Find images in a Writer document, OCR them, assign alt text and
(where necessary) inject a text equivalent block after the image.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import alt_text as altmod
from .ocr import OcrResult, run_ocr
from .uno_bridge import UnoBridge, path_to_url

log = logging.getLogger(__name__)


@dataclass
class ImageRecord:
    index: int
    name: str
    width_px: int
    height_px: int
    existing_alt: str
    assigned_alt: str
    decorative: bool
    ocr_text: str
    ocr_confidence: float
    ocr_usable: bool
    alt_confidence: float
    text_equivalent_injected: bool
    reasoning: str

    def as_dict(self) -> dict:
        return dict(self.__dict__)


@dataclass
class ImageStats:
    total: int = 0
    with_ocr: int = 0
    with_usable_ocr: int = 0
    with_existing_alt: int = 0
    newly_described: int = 0
    text_equivalents_added: int = 0
    decorative: int = 0
    records: list[ImageRecord] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "total": self.total,
            "with_ocr": self.with_ocr,
            "with_usable_ocr": self.with_usable_ocr,
            "with_existing_alt": self.with_existing_alt,
            "newly_described": self.newly_described,
            "text_equivalents_added": self.text_equivalents_added,
            "decorative": self.decorative,
            "records": [r.as_dict() for r in self.records],
        }


def _list_images(doc) -> list:
    out = []
    try:
        graphics = doc.getGraphicObjects()
    except Exception:
        graphics = None
    if graphics is not None:
        for i in range(graphics.getCount()):
            out.append(graphics.getByIndex(i))
    # DrawPage shapes (newer images may live here in some docs)
    try:
        page = doc.getDrawPage()
    except Exception:
        page = None
    if page is not None:
        for i in range(page.getCount()):
            shape = page.getByIndex(i)
            if shape.supportsService("com.sun.star.drawing.GraphicObjectShape") or \
               shape.supportsService("com.sun.star.text.TextGraphicObject"):
                if shape not in out:
                    out.append(shape)
    return out


def _export_image_bytes(ctx, shape, tmp_dir: Path, idx: int) -> tuple[bytes, int, int]:
    """Export a shape's graphic to a PNG and read its bytes."""
    graphic = None
    for attr in ("Graphic", "GraphicProvider", "GraphicURL"):
        try:
            graphic = shape.Graphic
            break
        except Exception:
            continue
    if graphic is None:
        return b"", 0, 0

    out_path = tmp_dir / f"img_{idx:03d}.png"
    smgr = ctx.ServiceManager
    provider = smgr.createInstanceWithContext(
        "com.sun.star.graphic.GraphicProvider", ctx
    )
    from com.sun.star.beans import PropertyValue  # type: ignore

    def pv(n, v):
        p = PropertyValue()
        p.Name = n
        p.Value = v
        return p

    url = path_to_url(out_path)
    store_props = (pv("URL", url), pv("MimeType", "image/png"))
    try:
        provider.storeGraphic(graphic, store_props)
    except Exception as e:
        log.warning("could not export image %d: %s", idx, e)
        return b"", 0, 0

    try:
        data = out_path.read_bytes()
    except FileNotFoundError:
        return b"", 0, 0

    w = h = 0
    try:
        from PIL import Image
        with Image.open(out_path) as im:
            w, h = im.size
    except Exception:
        pass
    return data, w, h


def _get_existing_alt(shape) -> str:
    for attr in ("Title", "Description"):
        try:
            v = getattr(shape, attr, "") or ""
            if v:
                return str(v)
        except Exception:
            continue
    return ""


def _set_alt(shape, alt: str, decorative: bool) -> None:
    try:
        if decorative:
            shape.Title = ""
            shape.Description = ""
        else:
            shape.Description = alt
            # Title is optional; keep short or leave existing
            if not getattr(shape, "Title", ""):
                shape.Title = alt[:80]
    except Exception as e:
        log.warning("setting alt failed: %s", e)


def _inject_equivalent(doc, shape, text: str) -> bool:
    """Insert a plain paragraph with the OCR'd text directly after the image's anchor."""
    if not text:
        return False
    try:
        anchor = shape.Anchor  # text cursor-like reference
    except Exception:
        anchor = None
    if anchor is None:
        return False
    try:
        text_ctx = anchor.getText()
        cursor = text_ctx.createTextCursorByRange(anchor.getEnd())
        # Break paragraph so new block is clearly separate
        text_ctx.insertControlCharacter(cursor, uno_getParagraphBreak(), False)
        text_ctx.insertString(cursor, "Текстовое содержимое изображения: " + text, False)
        try:
            cursor.ParaStyleName = "Quotations"
        except Exception:
            pass
        return True
    except Exception as e:
        log.warning("inject equivalent failed: %s", e)
        return False


def uno_getParagraphBreak() -> int:
    # com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK == 0
    return 0


def process_images(bridge: UnoBridge, doc, work_dir: Path) -> ImageStats:
    stats = ImageStats()
    shapes = _list_images(doc)
    stats.total = len(shapes)
    log.info("images found: %d", stats.total)
    if not shapes:
        return stats

    tmp_dir = work_dir / "images"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    for i, shape in enumerate(shapes):
        name = getattr(shape, "Name", f"img{i}") or f"img{i}"
        existing = _get_existing_alt(shape)
        if existing:
            stats.with_existing_alt += 1

        img_bytes, w, h = _export_image_bytes(bridge.ctx, shape, tmp_dir, i)
        if img_bytes:
            ocr = run_ocr(img_bytes)
            stats.with_ocr += 1
            if ocr.usable:
                stats.with_usable_ocr += 1
        else:
            ocr = OcrResult("", 0.0, "rus+eng", 0, False)

        decision = altmod.decide(ocr, w, h)
        _set_alt(shape, decision.alt_text, decision.decorative)
        if decision.decorative:
            stats.decorative += 1
        else:
            stats.newly_described += 1

        injected = False
        if decision.text_equivalent:
            injected = _inject_equivalent(doc, shape, decision.text_equivalent)
            if injected:
                stats.text_equivalents_added += 1

        stats.records.append(
            ImageRecord(
                index=i,
                name=name,
                width_px=w,
                height_px=h,
                existing_alt=existing,
                assigned_alt=decision.alt_text,
                decorative=decision.decorative,
                ocr_text=ocr.text,
                ocr_confidence=ocr.confidence,
                ocr_usable=ocr.usable,
                alt_confidence=decision.confidence,
                text_equivalent_injected=injected,
                reasoning=decision.reasoning,
            )
        )
    log.info("image stats: total=%d ocr=%d usable=%d", stats.total, stats.with_ocr, stats.with_usable_ocr)
    return stats
