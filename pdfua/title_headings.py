"""Title extraction and heading restoration for Writer documents.

The goal is to give the document real semantic structure so tagged PDF
export has something to work with. We do two things:

1. `ensure_title` — figure out a human-readable Title and write it into
   the document's DocumentProperties + a synchronized Heading 1 if the
   document has no top-level heading yet.
2. `restore_headings` — walk paragraphs and promote ones that visually
   look like headings (bold, short, centered, larger font) to
   `Heading 1` / `Heading 2` / `Heading 3`. Ones already tagged as a
   heading style are left alone.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

log = logging.getLogger(__name__)

HEADING_STYLES = {"Heading 1", "Heading 2", "Heading 3", "Title"}


@dataclass
class HeadingStats:
    title_source: str = ""  # "existing-heading", "first-paragraph", "filename"
    title: str = ""
    promoted_h1: int = 0
    promoted_h2: int = 0
    promoted_h3: int = 0
    paragraphs_examined: int = 0

    def as_dict(self) -> dict:
        return dict(self.__dict__)


def _para_info(para) -> dict:
    """Extract heuristic info from a paragraph."""
    text = para.getString().strip()
    info = {
        "text": text,
        "len": len(text),
        "style": "",
        "bold": False,
        "centered": False,
        "size": 0.0,
        "all_caps": False,
        "upper_letters": False,
    }
    try:
        info["style"] = para.ParaStyleName or ""
    except Exception:
        pass
    try:
        info["centered"] = para.ParaAdjust == 3  # CENTER
    except Exception:
        pass
    # Inspect first text portion for character properties
    try:
        enum = para.createEnumeration()
        if enum.hasMoreElements():
            portion = enum.nextElement()
            try:
                info["bold"] = portion.CharWeight >= 150
            except Exception:
                pass
            try:
                info["size"] = float(portion.CharHeight or 0)
            except Exception:
                pass
    except Exception:
        pass
    if text:
        letters = [c for c in text if c.isalpha()]
        if letters:
            info["all_caps"] = sum(1 for c in letters if c.isupper()) / len(letters) > 0.85
        info["upper_letters"] = any(c.isalpha() for c in text)
    return info


def _looks_like_heading(info: dict, median_size: float) -> int:
    """Return heading level 1..3 or 0 if not a heading."""
    t = info["text"]
    if not t or len(t) > 180:
        return 0
    if t.endswith((".", "!", "?", ":")) and len(t.split()) > 12:
        # Long sentence ending with punctuation — probably body text
        return 0
    score = 0
    if info["bold"]:
        score += 2
    if info["centered"]:
        score += 1
    if info["size"] and median_size and info["size"] >= median_size + 2:
        score += 2
    if info["size"] and median_size and info["size"] >= median_size + 5:
        score += 1
    if info["all_caps"] and 2 <= len(t.split()) <= 15:
        score += 2
    if len(t.split()) <= 10 and info["upper_letters"] and not t.endswith("."):
        score += 1
    # Numbered heading like "1." or "1.1"
    if re.match(r"^\d+(\.\d+){0,3}\.?\s+\S", t):
        score += 2
    if score >= 4:
        # Guess depth from size / numbering
        m = re.match(r"^(\d+(?:\.\d+){0,3})\.?\s+", t)
        if m:
            depth = m.group(1).count(".") + 1
            return min(depth, 3)
        if info["size"] and median_size and info["size"] >= median_size + 6:
            return 1
        if info["size"] and median_size and info["size"] >= median_size + 3:
            return 2
        return 2
    return 0


def _iter_top_paragraphs(doc):
    enum = doc.getText().createEnumeration()
    while enum.hasMoreElements():
        el = enum.nextElement()
        if el.supportsService("com.sun.star.text.Paragraph"):
            yield el


def _first_non_empty_paragraph(doc):
    for p in _iter_top_paragraphs(doc):
        if p.getString().strip():
            return p
    return None


def _derive_title_from_text(text: str) -> str:
    """Turn an arbitrary paragraph into a short title."""
    text = re.sub(r"\s+", " ", text).strip()
    # Drop trailing punctuation
    text = text.rstrip(" .,:;—-")
    # If it's a long sentence, take the first clause up to 80 chars.
    if len(text) <= 80:
        return text
    # Try to cut at sentence boundary
    m = re.match(r"(.{30,80}?)(?:[.!?]|,\s|:\s|—\s)", text)
    if m:
        return m.group(1).rstrip(" ,;:—-")
    return text[:80].rstrip(" ,;:—-") + "…"


def ensure_title(doc, source_path: Path) -> HeadingStats:
    stats = HeadingStats()
    title = ""
    title_source = ""

    # 1. Existing Heading 1 / Title style
    for para in _iter_top_paragraphs(doc):
        stats.paragraphs_examined += 1
        style = ""
        try:
            style = para.ParaStyleName or ""
        except Exception:
            pass
        text = para.getString().strip()
        if not text:
            continue
        if style in ("Title", "Heading 1"):
            title = text
            title_source = "existing-heading"
            break

    # 2. First non-empty paragraph
    if not title:
        p = _first_non_empty_paragraph(doc)
        if p is not None:
            title = _derive_title_from_text(p.getString())
            if title:
                title_source = "first-paragraph"

    # 3. Cleaned file name
    if not title:
        stem = source_path.stem
        stem = re.sub(r"[_\-]+", " ", stem).strip()
        title = stem or "Документ"
        title_source = "filename"

    # Write into document properties
    try:
        props = doc.getDocumentProperties()
        props.Title = title
        # Subject stays untouched
    except Exception as e:
        log.warning("could not set DocumentProperties.Title: %s", e)

    stats.title = title
    stats.title_source = title_source
    log.info("title=%r source=%s", title, title_source)
    return stats


def restore_headings(doc, stats: HeadingStats) -> HeadingStats:
    paras = list(_iter_top_paragraphs(doc))
    if not paras:
        return stats

    sizes = []
    for p in paras:
        try:
            enum = p.createEnumeration()
            if enum.hasMoreElements():
                portion = enum.nextElement()
                size = float(getattr(portion, "CharHeight", 0) or 0)
                if size > 0:
                    sizes.append(size)
        except Exception:
            continue
    median_size = sorted(sizes)[len(sizes) // 2] if sizes else 11.0

    first_heading_promoted = False
    for para in paras:
        info = _para_info(para)
        if info["style"] in HEADING_STYLES:
            first_heading_promoted = True
            continue
        level = _looks_like_heading(info, median_size)
        if not level:
            continue
        style = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3"}[level]
        try:
            para.ParaStyleName = style
            if level == 1:
                stats.promoted_h1 += 1
                first_heading_promoted = True
            elif level == 2:
                stats.promoted_h2 += 1
            else:
                stats.promoted_h3 += 1
        except Exception:
            continue

    # Guarantee there's at least one H1 — promote the first paragraph that
    # matches the document title to Heading 1 if nothing else qualified.
    if not first_heading_promoted:
        p = _first_non_empty_paragraph(doc)
        if p is not None:
            try:
                p.ParaStyleName = "Heading 1"
                stats.promoted_h1 += 1
            except Exception:
                pass

    return stats
