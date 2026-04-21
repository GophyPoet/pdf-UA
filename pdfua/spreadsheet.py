"""Spreadsheet path: open XLS/XLSX/ODS in Calc, clean up per-sheet
issues, then emit a fresh ODT with each sheet as a Heading 2 section
containing a proper Writer table.

Why not just convert XLSX → ODT with LibreOffice's direct filter?
Because the direct filter inherits every merge, empty column, narrow
column and layout quirk of the source workbook; the resulting ODT is
rarely accessible and often overflows page margins. We build a clean
ODT instead.

Pipeline:
  1. Open in Calc, iterate cells and pull every string value.
     Horizontally-merged headers leave empty cells with no value; we
     recover them in step 2 rather than trying to re-interpret Calc's
     merge metadata (which has historically been unreliable over UNO).
  2. Header-row recovery: for the top rows (before the first mostly-
     filled row), left-fill then down-fill empty cells. This restores
     the header semantics that merges encoded in the source.
  3. Drop columns that are entirely empty across all rows. Drop rows
     that are entirely empty.
  4. Truncate cell content that wouldn't fit in the target column
     width (wide Writer tables with 200-char-plus cells kill soffice's
     layout engine). Record the truncation count in the report.
  5. Page geometry is picked based on the widest sheet:
       ≤ 6 cols   → A4 portrait, default margins
       7..10 cols → A4 landscape
       ≥ 11 cols  → A4 landscape with narrow margins + smaller font
  6. Set explicit relative column widths summing to 65535 (Writer's
     unit for TableColumnRelativeSums) and pin table width to the
     usable page width, so nothing escapes the PDF page box.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

from .uno_bridge import UnoBridge

log = logging.getLogger(__name__)


# All dimensions below are in 1/100 mm (soffice's internal unit).
A4_LONG = 29700
A4_SHORT = 21000


@dataclass
class SheetReport:
    name: str
    rows: int
    cols: int
    empty_rows_skipped: int = 0
    empty_cols_skipped: int = 0
    header_cells_filled: int = 0
    cells_truncated: int = 0
    cells_filled: int = 0
    skipped: bool = False

    def as_dict(self) -> dict:
        return dict(self.__dict__)


@dataclass
class SpreadsheetStats:
    sheets: list[SheetReport] = field(default_factory=list)
    total_sheets: int = 0
    total_header_cells_filled: int = 0
    total_cells_truncated: int = 0
    total_empty_cells_filled: int = 0
    total_empty_rows_skipped: int = 0
    total_empty_cols_skipped: int = 0
    page_orientation: str = "portrait"
    max_cols: int = 0

    def as_dict(self) -> dict:
        return {
            "total_sheets": self.total_sheets,
            "total_header_cells_filled": self.total_header_cells_filled,
            "total_cells_truncated": self.total_cells_truncated,
            "total_empty_cells_filled": self.total_empty_cells_filled,
            "total_empty_rows_skipped": self.total_empty_rows_skipped,
            "total_empty_cols_skipped": self.total_empty_cols_skipped,
            "page_orientation": self.page_orientation,
            "max_cols": self.max_cols,
            "sheets": [s.as_dict() for s in self.sheets],
        }


def _sheet_used_range(sheet) -> tuple[int, int, int, int]:
    try:
        cursor = sheet.createCursor()
        cursor.gotoStartOfUsedArea(False)
        cursor.gotoEndOfUsedArea(True)
        addr = cursor.getRangeAddress()
        return addr.StartRow, addr.StartColumn, addr.EndRow, addr.EndColumn
    except Exception:
        return 0, 0, 0, 0


def _extract_raw_matrix(sheet) -> list[list[str]]:
    """Pull a rectangular string matrix of the used area, preserving
    empty cells (they matter for header-row recovery)."""
    r0, c0, r1, c1 = _sheet_used_range(sheet)
    if (r0, c0, r1, c1) == (0, 0, 0, 0):
        return []
    out: list[list[str]] = []
    for r in range(r0, r1 + 1):
        row: list[str] = []
        for c in range(c0, c1 + 1):
            try:
                v = sheet.getCellByPosition(c, r).getString()
            except Exception:
                v = ""
            row.append((v or "").replace("\xa0", " ").strip())
        out.append(row)
    return out


def _header_row_count(matrix: list[list[str]]) -> int:
    """Heuristic: the header block is the leading run of rows that are
    less than 85% filled AND less filled than the first 'data' row.
    Usually 1 — sometimes 2..4 for plan-type sheets."""
    if not matrix:
        return 0
    cols = max(len(r) for r in matrix)
    # First 'data' row = first row with fill ratio >= 0.85
    first_data = None
    for i, row in enumerate(matrix):
        filled = sum(1 for v in row if v)
        if cols and filled / cols >= 0.85:
            first_data = i
            break
    if first_data is None:
        return 1  # no clearly-populated data row; treat first row as header
    return max(1, first_data)


def _recover_header_rows(matrix: list[list[str]], header_rows: int) -> int:
    """For rows 0..header_rows-1, left-fill empty cells from the nearest
    non-empty cell to their left, then down-fill within the header block
    from the nearest non-empty cell above. Returns count of cells
    populated by this recovery."""
    filled = 0
    if not matrix or header_rows <= 0:
        return 0
    cols = max(len(r) for r in matrix)
    # Normalize row widths.
    for row in matrix:
        while len(row) < cols:
            row.append("")

    # Left-fill header rows.
    for r in range(min(header_rows, len(matrix))):
        last = ""
        for c in range(cols):
            v = matrix[r][c]
            if v:
                last = v
            elif last:
                matrix[r][c] = last
                filled += 1

    # Down-fill within header block: for each column, carry forward
    # non-empty values down.
    for c in range(cols):
        last = ""
        for r in range(min(header_rows, len(matrix))):
            v = matrix[r][c]
            if v:
                last = v
            elif last:
                matrix[r][c] = last
                filled += 1
    return filled


def _collapse_multirow_header(matrix: list[list[str]], header_rows: int) -> tuple[list[str], list[list[str]]]:
    """Collapse the multi-row header block into a single header row.
    For each column, join the non-empty header values with ' / ' and
    dedupe consecutive duplicates. Returns (header_row, data_rows).
    """
    if not matrix:
        return [], []
    cols = max(len(r) for r in matrix)
    for row in matrix:
        while len(row) < cols:
            row.append("")
    header_block = matrix[:header_rows]
    data_rows = matrix[header_rows:]
    header: list[str] = []
    for c in range(cols):
        parts: list[str] = []
        for r in range(len(header_block)):
            v = header_block[r][c].strip()
            if v and (not parts or parts[-1] != v):
                parts.append(v)
        header.append(" / ".join(parts))
    return header, data_rows


def _drop_empty(matrix: list[list[str]], report: SheetReport) -> list[list[str]]:
    """Drop entirely-empty columns and rows. Fill remaining empties with —."""
    if not matrix:
        return []
    cols = max(len(r) for r in matrix)
    keep_cols = [
        c for c in range(cols) if any((row[c] if c < len(row) else "") for row in matrix)
    ]
    report.empty_cols_skipped = cols - len(keep_cols)
    matrix = [[row[c] if c < len(row) else "" for c in keep_cols] for row in matrix]

    before = len(matrix)
    matrix = [row for row in matrix if any(v for v in row)]
    report.empty_rows_skipped = before - len(matrix)

    for row in matrix:
        for i, v in enumerate(row):
            if not v:
                row[i] = "—"
                report.cells_filled += 1
    return matrix


def _truncate_for_width(matrix: list[list[str]], cols: int, usable_hmm: int, font_pt: int, report: SheetReport) -> None:
    """Truncate cell content so a single cell never exceeds what LibreOffice
    can safely lay out at the configured column width. Strings longer than
    roughly 12 wrapped lines of text at the given column width blow up the
    table layout engine.

    We compute a hard max-chars-per-cell from column width and font size,
    then truncate with an ellipsis marker '…'.
    """
    if not matrix or cols == 0:
        return
    per_col_hmm = usable_hmm / cols
    # Approx char width at 8pt ≈ 1.5mm, at 9pt ≈ 1.7mm, at 10pt ≈ 1.9mm.
    char_mm = {8: 1.5, 9: 1.7, 10: 1.9}.get(font_pt, 1.7)
    chars_per_line = max(3, int(per_col_hmm / 100 / char_mm))
    # Cap at a max number of wrapped lines per cell so the row stays within
    # a single printed page. 8 lines at 8pt ≈ 32mm — comfortable.
    max_chars = chars_per_line * 8
    # Always keep at least 40 chars available so short-column tables don't
    # amputate useful content.
    max_chars = max(40, max_chars)

    for row in matrix:
        for i, v in enumerate(row):
            if len(v) > max_chars:
                row[i] = v[: max_chars - 1] + "…"
                report.cells_truncated += 1


def _configure_page(writer_doc, max_cols: int) -> tuple[str, int, int]:
    """Pick page orientation + usable width based on the widest sheet.

    Returns (orientation, usable_width_hmm, font_size_pt).
    """
    try:
        styles = writer_doc.getStyleFamilies().getByName("PageStyles")
        style = styles.getByName("Default Page Style")
    except Exception:
        try:
            style = styles.getByName("Standard")
        except Exception:
            return "portrait", 17000, 10

    if max_cols <= 6:
        try:
            style.setPropertyValue("IsLandscape", False)
            style.setPropertyValue("Width", A4_SHORT)
            style.setPropertyValue("Height", A4_LONG)
            style.setPropertyValue("LeftMargin", 2000)
            style.setPropertyValue("RightMargin", 2000)
            style.setPropertyValue("TopMargin", 2000)
            style.setPropertyValue("BottomMargin", 2000)
        except Exception:
            pass
        return "portrait", A4_SHORT - 4000, 10

    if max_cols <= 10:
        try:
            style.setPropertyValue("IsLandscape", True)
            style.setPropertyValue("Width", A4_LONG)
            style.setPropertyValue("Height", A4_SHORT)
            style.setPropertyValue("LeftMargin", 1500)
            style.setPropertyValue("RightMargin", 1500)
            style.setPropertyValue("TopMargin", 1500)
            style.setPropertyValue("BottomMargin", 1500)
        except Exception:
            pass
        return "landscape", A4_LONG - 3000, 9

    try:
        style.setPropertyValue("IsLandscape", True)
        style.setPropertyValue("Width", A4_LONG)
        style.setPropertyValue("Height", A4_SHORT)
        style.setPropertyValue("LeftMargin", 800)
        style.setPropertyValue("RightMargin", 800)
        style.setPropertyValue("TopMargin", 1000)
        style.setPropertyValue("BottomMargin", 1000)
    except Exception:
        pass
    return "landscape-narrow", A4_LONG - 1600, 8


def _relative_column_widths(matrix: list[list[str]]) -> list[int]:
    """Weights proportional to max string length per column, floored."""
    if not matrix:
        return []
    cols = max(len(r) for r in matrix)
    weights = []
    for c in range(cols):
        maxlen = 4
        for row in matrix:
            if c < len(row):
                maxlen = max(maxlen, min(len(row[c]), 40))
        weights.append(maxlen)
    total = sum(weights)
    # Convert to cumulative relative sums scaled to 65535 (Writer's unit).
    cum = []
    running = 0
    for w in weights:
        running += w
        cum.append(int(running * 65535 / total))
    return cum


def _insert_writer_table(
    writer_doc,
    matrix: list[list[str]],
    heading: str,
    usable_hmm: int,
    font_pt: int,
) -> None:
    text = writer_doc.getText()
    cursor = text.createTextCursorByRange(text.getEnd())

    if heading:
        try:
            cursor.ParaStyleName = "Heading 2"
        except Exception:
            pass
        text.insertString(cursor, heading, False)
        text.insertControlCharacter(cursor, 0, False)
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

    table = writer_doc.createInstance("com.sun.star.text.TextTable")
    table.initialize(rows, cols)
    text.insertTextContent(cursor, table, False)

    # Bulk set values via DataArray — faster and more robust than
    # per-cell setString for wide tables.
    try:
        data = tuple(
            tuple((row[c] if c < len(row) else "") for c in range(cols))
            for row in matrix
        )
        table.getCellRangeByPosition(0, 0, cols - 1, rows - 1).setDataArray(data)
    except Exception:
        letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        def col_name(i: int) -> str:
            s, i2 = "", i
            while True:
                s = letters[i2 % 26] + s
                i2 = i2 // 26 - 1
                if i2 < 0:
                    break
            return s
        for r, row in enumerate(matrix):
            for c in range(cols):
                try:
                    table.getCellByName(f"{col_name(c)}{r + 1}").setString(
                        row[c] if c < len(row) else ""
                    )
                except Exception:
                    continue

    # Pin table to usable page width.
    try:
        table.setPropertyValue("RelativeWidth", 100)
        table.setPropertyValue("Width", usable_hmm)
        table.setPropertyValue("HoriOrient", 3)  # LEFT_AND_WIDTH
    except Exception:
        pass

    # Column widths proportional to content length.
    try:
        cum = _relative_column_widths(matrix)
        if cum:
            table.setPropertyValue("TableColumnRelativeSums", tuple(cum))
    except Exception:
        pass

    try:
        table.RepeatHeadline = True
        table.HeaderRowCount = 1
    except Exception:
        pass

    try:
        rng = table.getCellRangeByPosition(0, 0, cols - 1, rows - 1)
        rng.CharHeight = float(font_pt)
        rng.CharHeightAsian = float(font_pt)
        rng.CharHeightComplex = float(font_pt)
    except Exception:
        pass
    try:
        header = table.getCellRangeByPosition(0, 0, cols - 1, 0)
        header.CharWeight = 150.0
        header.CharWeightAsian = 150.0
    except Exception:
        pass

    cursor = text.createTextCursorByRange(text.getEnd())
    text.insertControlCharacter(cursor, 0, False)


def _prepare_sheet(sheet, rep: SheetReport) -> list[list[str]]:
    raw = _extract_raw_matrix(sheet)
    if not raw:
        return []
    hdr_rows = _header_row_count(raw)
    filled = _recover_header_rows(raw, hdr_rows)
    rep.header_cells_filled = filled
    # Collapse multi-row header into a single row. Sub-header rows 1..N-1
    # are dropped; the data rows remain.
    header, data_rows = _collapse_multirow_header(raw, hdr_rows)
    matrix = [header] + data_rows if header else data_rows
    # Drop empties + fill with —.
    matrix = _drop_empty(matrix, rep)
    rep.rows = len(matrix)
    rep.cols = len(matrix[0]) if matrix else 0
    return matrix


def build_odt_from_spreadsheet(
    bridge: UnoBridge, source: Path, dest_odt: Path, title: str
) -> SpreadsheetStats:
    calc_doc = bridge.load(source, hidden=True)
    stats = SpreadsheetStats()
    try:
        sheets = calc_doc.getSheets()
        stats.total_sheets = sheets.getCount()

        # Pass 1: extract every sheet to a clean matrix.
        per_sheet: list[tuple[SheetReport, list[list[str]]]] = []
        for i in range(stats.total_sheets):
            sheet = sheets.getByIndex(i)
            rep = SheetReport(name=sheet.getName(), rows=0, cols=0)
            try:
                matrix = _prepare_sheet(sheet, rep)
            except Exception as e:
                log.warning("sheet %s: extraction failed: %s", rep.name, e)
                matrix = []
                rep.skipped = True
            per_sheet.append((rep, matrix))
            if matrix:
                stats.max_cols = max(stats.max_cols, len(matrix[0]))
            stats.total_header_cells_filled += rep.header_cells_filled
            stats.total_empty_cells_filled += rep.cells_filled
            stats.total_empty_rows_skipped += rep.empty_rows_skipped
            stats.total_empty_cols_skipped += rep.empty_cols_skipped

        # We don't need the calc doc any more — close it before touching
        # Writer to reduce the chance of UNO bridge hiccups mixing two
        # open components.
        try:
            calc_doc.close(True)
            calc_doc = None
        except Exception:
            pass

        writer_doc = bridge.new_writer()
        try:
            try:
                writer_doc.getDocumentProperties().Title = title
            except Exception:
                pass

            orientation, usable_hmm, font_pt = _configure_page(writer_doc, stats.max_cols)
            stats.page_orientation = orientation
            log.info(
                "page: %s, usable_width=%d(1/100mm), font=%dpt, max_cols=%d",
                orientation, usable_hmm, font_pt, stats.max_cols,
            )

            # Pass 2: now that we know the final geometry, truncate any
            # cell content that would blow the column width / row-height
            # layout limit.
            for rep, matrix in per_sheet:
                if matrix:
                    _truncate_for_width(matrix, len(matrix[0]), usable_hmm, font_pt, rep)
                    stats.total_cells_truncated += rep.cells_truncated

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

            for rep, matrix in per_sheet:
                if not matrix:
                    rep.skipped = True
                    stats.sheets.append(rep)
                    continue
                try:
                    _insert_writer_table(writer_doc, matrix, rep.name, usable_hmm, font_pt)
                except Exception as e:
                    log.warning("sheet %s: insert failed: %s", rep.name, e)
                    rep.skipped = True
                stats.sheets.append(rep)

            bridge.save_as(writer_doc, dest_odt, "writer8")
        finally:
            try:
                writer_doc.close(True)
            except Exception:
                pass
    finally:
        if calc_doc is not None:
            try:
                calc_doc.close(True)
            except Exception:
                pass
    return stats
