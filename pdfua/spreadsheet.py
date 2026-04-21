"""Spreadsheet path: open XLS/XLSX/ODS in Calc, clean up per-sheet
issues, then emit a fresh ODT with each sheet as a Heading 2 section
containing a proper Writer table.

Why not just convert XLSX → ODT with LibreOffice? Because the direct
filter inherits all the merges / empty columns / layout quirks of the
source workbook and the resulting ODT is rarely accessible. We build
a clean ODT instead.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .uno_bridge import UnoBridge, path_to_url, props

log = logging.getLogger(__name__)


@dataclass
class SheetReport:
    name: str
    rows: int
    cols: int
    empty_rows_skipped: int = 0
    empty_cols_skipped: int = 0
    merged_regions: int = 0
    cells_filled: int = 0

    def as_dict(self) -> dict:
        return dict(self.__dict__)


@dataclass
class SpreadsheetStats:
    sheets: list[SheetReport] = field(default_factory=list)
    total_sheets: int = 0
    total_merged_regions: int = 0
    total_empty_cells_filled: int = 0
    total_empty_rows_skipped: int = 0
    total_empty_cols_skipped: int = 0

    def as_dict(self) -> dict:
        return {
            "total_sheets": self.total_sheets,
            "total_merged_regions": self.total_merged_regions,
            "total_empty_cells_filled": self.total_empty_cells_filled,
            "total_empty_rows_skipped": self.total_empty_rows_skipped,
            "total_empty_cols_skipped": self.total_empty_cols_skipped,
            "sheets": [s.as_dict() for s in self.sheets],
        }


def _sheet_used_range(sheet) -> tuple[int, int, int, int]:
    """Return (first_row, first_col, last_row, last_col) for the used area."""
    try:
        cursor = sheet.createCursor()
        cursor.gotoStartOfUsedArea(False)
        cursor.gotoEndOfUsedArea(True)
        addr = cursor.getRangeAddress()
        return addr.StartRow, addr.StartColumn, addr.EndRow, addr.EndColumn
    except Exception:
        return 0, 0, 0, 0


def _unmerge_all(sheet, report: SheetReport) -> None:
    try:
        cursor = sheet.createCursor()
        cursor.gotoStartOfUsedArea(False)
        cursor.gotoEndOfUsedArea(True)
    except Exception:
        return
    try:
        addr = cursor.getRangeAddress()
        r0, c0, r1, c1 = addr.StartRow, addr.StartColumn, addr.EndRow, addr.EndColumn
    except Exception:
        return
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            try:
                cell = sheet.getCellByPosition(c, r)
                if cell.getIsMerged():
                    # find the merged range and unmerge it
                    range_obj = sheet.getCellRangeByPosition(c, r, c, r).queryIntersection(
                        cell.getSpreadsheet().getRangeAddress()
                    )
            except Exception:
                continue
    # Simpler and more reliable approach: iterate merged ranges if exposed.
    try:
        # Calc exposes MergedRanges via each cell range; we iterate known region
        region = sheet.getCellRangeByPosition(c0, r0, c1, r1)
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                try:
                    cr = sheet.getCellRangeByPosition(c, r, c, r)
                    if cr.getIsMerged():
                        cr.merge(False)
                        report.merged_regions += 1
                except Exception:
                    continue
    except Exception:
        pass


def _extract_sheet_matrix(sheet, report: SheetReport) -> list[list[str]]:
    r0, c0, r1, c1 = _sheet_used_range(sheet)
    if (r0, c0, r1, c1) == (0, 0, 0, 0):
        # empty sheet
        return []
    raw: list[list[str]] = []
    for r in range(r0, r1 + 1):
        row_vals: list[str] = []
        for c in range(c0, c1 + 1):
            try:
                cell = sheet.getCellByPosition(c, r)
                v = cell.getString()
            except Exception:
                v = ""
            row_vals.append((v or "").strip())
        raw.append(row_vals)
    report.rows = len(raw)
    report.cols = len(raw[0]) if raw else 0

    # Drop columns that are entirely empty across all rows
    if raw:
        keep_cols = [
            c for c in range(len(raw[0])) if any(row[c] for row in raw if c < len(row))
        ]
        report.empty_cols_skipped = report.cols - len(keep_cols)
        raw = [[row[c] if c < len(row) else "" for c in keep_cols] for row in raw]

    # Drop rows that are entirely empty
    before = len(raw)
    raw = [row for row in raw if any(v for v in row)]
    report.empty_rows_skipped = before - len(raw)

    # Fill remaining empty cells with placeholder
    for row in raw:
        for i, v in enumerate(row):
            if not v:
                row[i] = "—"
                report.cells_filled += 1
    report.cols = len(raw[0]) if raw else 0
    report.rows = len(raw)
    return raw


def _insert_writer_table(writer_doc, matrix: list[list[str]], heading: str) -> None:
    text = writer_doc.getText()
    cursor = text.createTextCursorByRange(text.getEnd())

    # Heading
    if heading:
        try:
            cursor.ParaStyleName = "Heading 2"
        except Exception:
            pass
        text.insertString(cursor, heading, False)
        text.insertControlCharacter(cursor, 0, False)  # paragraph break
        try:
            cursor.ParaStyleName = "Default Paragraph Style"
        except Exception:
            pass

    if not matrix:
        text.insertString(cursor, "(лист не содержит данных)", False)
        text.insertControlCharacter(cursor, 0, False)
        return

    rows = len(matrix)
    cols = max(len(r) for r in matrix)

    smgr = writer_doc.getInternalServiceManager() if hasattr(writer_doc, "getInternalServiceManager") else None
    # Create the table via the document factory
    table = writer_doc.createInstance("com.sun.star.text.TextTable")
    table.initialize(rows, cols)
    text.insertTextContent(cursor, table, False)

    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    def col_name(i: int) -> str:
        s = ""
        i2 = i
        while True:
            s = letters[i2 % 26] + s
            i2 = i2 // 26 - 1
            if i2 < 0:
                break
        return s

    for r, row in enumerate(matrix):
        for c in range(cols):
            v = row[c] if c < len(row) else ""
            cname = f"{col_name(c)}{r + 1}"
            try:
                cell = table.getCellByName(cname)
                cell.setString(v or "")
            except Exception:
                continue

    # Header row
    try:
        table.RepeatHeadline = True
        table.HeaderRowCount = 1
    except Exception:
        pass

    # Paragraph break after table so the next sheet's heading doesn't
    # merge into this table.
    cursor = text.createTextCursorByRange(text.getEnd())
    text.insertControlCharacter(cursor, 0, False)


def build_odt_from_spreadsheet(
    bridge: UnoBridge, source: Path, dest_odt: Path, title: str
) -> SpreadsheetStats:
    calc_doc = bridge.load(source, hidden=True)
    stats = SpreadsheetStats()
    try:
        sheets = calc_doc.getSheets()
        writer_doc = bridge.new_writer()
        try:
            # Doc title + top-level heading
            try:
                writer_doc.getDocumentProperties().Title = title
            except Exception:
                pass
            text = writer_doc.getText()
            cursor = text.createTextCursorByRange(text.getEnd())
            try:
                cursor.ParaStyleName = "Heading 1"
            except Exception:
                pass
            text.insertString(cursor, title or "Документ", False)
            text.insertControlCharacter(cursor, 0, False)
            try:
                cursor.ParaStyleName = "Default Paragraph Style"
            except Exception:
                pass

            stats.total_sheets = sheets.getCount()
            for i in range(stats.total_sheets):
                sheet = sheets.getByIndex(i)
                report = SheetReport(name=sheet.getName(), rows=0, cols=0)
                _unmerge_all(sheet, report)
                matrix = _extract_sheet_matrix(sheet, report)
                _insert_writer_table(writer_doc, matrix, sheet.getName())
                stats.sheets.append(report)
                stats.total_merged_regions += report.merged_regions
                stats.total_empty_cells_filled += report.cells_filled
                stats.total_empty_rows_skipped += report.empty_rows_skipped
                stats.total_empty_cols_skipped += report.empty_cols_skipped

            # Save intermediate ODT
            bridge.save_as(writer_doc, dest_odt, "writer8")
        finally:
            writer_doc.close(True)
    finally:
        calc_doc.close(True)
    return stats
