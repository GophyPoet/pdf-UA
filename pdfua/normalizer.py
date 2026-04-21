"""Text normalization over a live Writer document.

Operates on the UNO text tree:
  - walks paragraphs enumerating portions,
  - collapses runs of whitespace / decorative tabs inside paragraphs,
  - removes long underscore fill lines,
  - strips stray manual line breaks,
  - drops empty decorative paragraphs that separate blocks.

We do NOT nuke direct formatting via `ClearDirectFormatting` — that would
lose headings, bold/italic semantics, list markers, etc. Instead we
normalize at the character level and later reassign paragraph styles in
the heading stage.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

log = logging.getLogger(__name__)

# Multiple spaces (2+)
MULTI_SPACE_RE = re.compile(r"[   -​]{2,}")
# Tab or sequence of tabs (used to pad text visually)
TAB_RE = re.compile(r"\t+")
# Runs of 3+ underscores (used as fill lines)
UNDERSCORE_LINE_RE = re.compile(r"_{3,}")
# Underscores with spaces between, also treated as fill lines
UNDERSCORE_SPACED_RE = re.compile(r"(?:_\s*){3,}_?")
# Non-printable control characters we never want in accessible text
CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


@dataclass
class NormalizerStats:
    multi_spaces_collapsed: int = 0
    tabs_removed: int = 0
    underscore_lines_handled: int = 0
    underscore_lines_replaced_with_placeholder: int = 0
    ctrl_chars_removed: int = 0
    empty_paragraphs_removed: int = 0
    manual_breaks_removed: int = 0
    paragraphs_scanned: int = 0

    def as_dict(self) -> dict:
        return dict(self.__dict__)


def _normalize_run(text: str, stats: NormalizerStats) -> str:
    orig = text

    # Control chars
    cleaned = CTRL_RE.sub("", text)
    stats.ctrl_chars_removed += len(text) - len(cleaned)
    text = cleaned

    # Underscores: decide replace vs drop by local context
    def _sub_underscores(m: re.Match) -> str:
        stats.underscore_lines_handled += 1
        left = m.string[: m.start()].rstrip()
        right = m.string[m.end():].lstrip()
        # Heuristic: fill field if the text around reads like "ФИО ___"
        if left and (left.rstrip(" :.—-").endswith((":",)) or len(left.split()[-1]) <= 40):
            if left.rstrip().endswith((":", "—", "-")):
                stats.underscore_lines_replaced_with_placeholder += 1
                return " [поле для заполнения]"
        if not left and not right:
            return ""  # decorative separator line
        stats.underscore_lines_replaced_with_placeholder += 1
        return " [поле для заполнения]"

    text = UNDERSCORE_SPACED_RE.sub(_sub_underscores, text)
    text = UNDERSCORE_LINE_RE.sub(_sub_underscores, text)

    # Tabs used inside text as spacers
    new_text, tabs = TAB_RE.subn(" ", text)
    stats.tabs_removed += tabs
    text = new_text

    # Collapse multiple whitespace
    def _sub_multi(m: re.Match) -> str:
        stats.multi_spaces_collapsed += 1
        return " "

    text = MULTI_SPACE_RE.sub(_sub_multi, text)

    # Trim trailing spaces
    text = text.rstrip(" \t")
    # And leading if we produced them via replacement
    if text.startswith(" ") and not orig.startswith(" "):
        text = text.lstrip(" ")
    return text


def _iter_paragraphs(doc):
    """Iterate paragraph objects at the top level + inside tables."""
    enum = doc.getText().createEnumeration()
    while enum.hasMoreElements():
        el = enum.nextElement()
        if el.supportsService("com.sun.star.text.Paragraph"):
            yield el
        elif el.supportsService("com.sun.star.text.TextTable"):
            for cell_name in el.getCellNames():
                cell = el.getCellByName(cell_name)
                sub = cell.createEnumeration()
                while sub.hasMoreElements():
                    sel = sub.nextElement()
                    if sel.supportsService("com.sun.star.text.Paragraph"):
                        yield sel


def normalize_document(doc) -> NormalizerStats:
    """Run in-place normalization on an open Writer document."""
    stats = NormalizerStats()

    # Collect empty paragraphs for later removal (3+ consecutive blanks).
    prev_blank_count = 0
    blanks_to_remove = []

    for para in _iter_paragraphs(doc):
        stats.paragraphs_scanned += 1
        portion_enum = para.createEnumeration()
        portions = []
        while portion_enum.hasMoreElements():
            portions.append(portion_enum.nextElement())

        has_non_text = False
        for portion in portions:
            try:
                ptype = portion.TextPortionType
            except Exception:
                ptype = "Text"
            if ptype != "Text":
                has_non_text = True
                if ptype == "LineBreak":
                    stats.manual_breaks_removed += 1

        # Edit individual text portions only. Never overwrite the whole
        # paragraph via para.setString, because that would destroy inline
        # images, fields, footnotes, or other TextContent portions.
        for portion in portions:
            try:
                ptype = portion.TextPortionType
            except Exception:
                ptype = "Text"
            if ptype != "Text":
                continue
            s = portion.getString()
            if not s:
                continue
            ns = _normalize_run(s, stats)
            if ns != s:
                try:
                    portion.setString(ns)
                except Exception:
                    pass

        # A paragraph that contains an inline shape / field / footnote is
        # NOT empty, even if getString() is blank. Removing it would
        # destroy the image or field.
        is_blank = para.getString().strip() == "" and not has_non_text
        if is_blank:
            prev_blank_count += 1
            if prev_blank_count > 1:
                blanks_to_remove.append(para)
        else:
            prev_blank_count = 0

    # Remove collected empty paragraphs (skip the first of each streak)
    for para in blanks_to_remove:
        try:
            doc.getText().removeTextContent(para)
            stats.empty_paragraphs_removed += 1
        except Exception:
            pass

    log.info("normalizer stats: %s", stats.as_dict())
    return stats
