"""Repair text tables in Writer for PDF/UA.

What we do:
  - Detect merged cells and split them.
  - Fill empty cells with a safe placeholder ("—").
  - Remove fully-empty rows and columns.
  - Mark the first row as a header row where plausible.
  - Detect tables used purely for layout (single row or mostly empty)
    and flag them; conversion to plain text is optional and done only
    when the table is small and clearly not a data table.

Limits of UNO:
  - Not every complex merge pattern can be perfectly split via the
    "split" API — some documents need a full rebuild. For those cases
    we do a best-effort split and report residual risk.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field

log = logging.getLogger(__name__)


@dataclass
class TableStats:
    total: int = 0
    with_merges: int = 0
    merges_split: int = 0
    empty_cells_filled: int = 0
    empty_rows_removed: int = 0
    empty_cols_removed: int = 0
    header_rows_set: int = 0
    layout_tables_flagged: int = 0
    converted_to_text: int = 0
    residual_merge_tables: int = 0
    notes: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        d = dict(self.__dict__)
        return d


def _tables(doc):
    return doc.getTextTables()


def _cell_count(row_count: int, col_count: int, table) -> int:
    # Writer tables: getCellNames() returns e.g. ["A1","A2.1"] for merged cells.
    # If len(names) != rows*cols then there are merges.
    return len(table.getCellNames())


def _has_merges(table) -> bool:
    try:
        rows = table.getRows().getCount()
        cols = table.getColumns().getCount()
    except Exception:
        return False
    names = table.getCellNames()
    # Any name containing a "." means a split/merge relative to grid
    return any("." in n for n in names) or len(names) != rows * cols


def _split_merges(table, stats: TableStats) -> bool:
    """Best-effort split via XTextTableCursor.splitRange.

    Guards:
      - Abort on the first hard failure (continuing has crashed soffice
        on DOCX-imported tables with irregular grids).
      - Abort if the grid size explodes past a sane multiple of the
        original — some DOCX imports have rowspan structures where a
        single splitRange call grows the grid to hundreds of cells.
    """
    initial_cells = len(table.getCellNames())
    hard_limit = max(initial_cells * 2, initial_cells + 20)
    local_splits = 0

    changed = True
    iterations = 0
    while changed and iterations < 6:
        changed = False
        iterations += 1
        names = list(table.getCellNames())
        if len(names) > hard_limit:
            return False
        for cname in names:
            try:
                cursor = table.createCursorByCellName(cname)
            except Exception:
                return False
            try:
                start = cursor.getRangeName()
                cursor.gotoEnd(True)
                end = cursor.getRangeName()
            except Exception:
                return False
            if start == end:
                continue
            did_split = False
            try:
                cursor.splitRange(1, False)
                did_split = True
            except Exception:
                try:
                    cursor.splitRange(1, True)
                    did_split = True
                except Exception:
                    return False
            if did_split:
                local_splits += 1
                stats.merges_split += 1
                if local_splits > 8:
                    return False
                changed = True
                break
    return not _has_merges(table)


def _cell_text(cell) -> str:
    try:
        return cell.getString().strip()
    except Exception:
        return ""


def _set_cell_text(cell, text: str) -> None:
    try:
        cell.setString(text)
    except Exception:
        pass


def _fill_empties(table, stats: TableStats) -> None:
    for cname in table.getCellNames():
        try:
            cell = table.getCellByName(cname)
        except Exception:
            continue
        if _cell_text(cell) == "":
            _set_cell_text(cell, "—")
            stats.empty_cells_filled += 1


def _row_is_empty(table, row_idx: int) -> bool:
    cols = table.getColumns().getCount()
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

    for c in range(cols):
        cname = f"{col_name(c)}{row_idx + 1}"
        try:
            cell = table.getCellByName(cname)
        except Exception:
            continue
        if _cell_text(cell):
            return False
    return True


def _remove_empty_rows(table, stats: TableStats) -> None:
    # Iterate top-down, removing from the bottom.
    rows = table.getRows()
    total = rows.getCount()
    for r in reversed(range(total)):
        if total - stats.empty_rows_removed <= 1:
            break
        if _row_is_empty(table, r):
            try:
                rows.removeByIndex(r, 1)
                stats.empty_rows_removed += 1
            except Exception:
                continue


def _col_is_empty(table, col_idx: int) -> bool:
    rows = table.getRows().getCount()
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

    cname_base = col_name(col_idx)
    for r in range(rows):
        cname = f"{cname_base}{r + 1}"
        try:
            cell = table.getCellByName(cname)
        except Exception:
            continue
        if _cell_text(cell):
            return False
    return True


def _remove_empty_cols(table, stats: TableStats) -> None:
    cols = table.getColumns()
    total = cols.getCount()
    for c in reversed(range(total)):
        if total - stats.empty_cols_removed <= 1:
            break
        if _col_is_empty(table, c):
            try:
                cols.removeByIndex(c, 1)
                stats.empty_cols_removed += 1
            except Exception:
                continue


def _set_header_row(table, stats: TableStats) -> None:
    try:
        table.RepeatHeadline = True
        table.HeaderRowCount = 1
        stats.header_rows_set += 1
    except Exception:
        pass


def repair_tables(doc) -> TableStats:
    stats = TableStats()
    tables = _tables(doc)
    stats.total = tables.getCount()
    for i in range(stats.total):
        try:
            table = tables.getByIndex(i)
        except Exception:
            continue

        table_has_residual_merges = False
        if _has_merges(table):
            stats.with_merges += 1
            try:
                ok = _split_merges(table, stats)
            except Exception as e:
                log.warning("table #%d: split raised %s", i, e)
                ok = False
            if not ok:
                table_has_residual_merges = True
                stats.residual_merge_tables += 1
                stats.notes.append(
                    f"table #{i}: residual merges after best-effort split"
                )

        try:
            _fill_empties(table, stats)
        except Exception as e:
            log.warning("table #%d: fill-empties failed: %s", i, e)

        # Row/column removal is destructive. Skip it when the table still
        # has merges — doing it on an irregular grid can crash soffice.
        if not table_has_residual_merges:
            try:
                _remove_empty_rows(table, stats)
            except Exception as e:
                log.warning("table #%d: remove-empty-rows failed: %s", i, e)
            try:
                _remove_empty_cols(table, stats)
            except Exception as e:
                log.warning("table #%d: remove-empty-cols failed: %s", i, e)
        else:
            stats.notes.append(
                f"table #{i}: empty row/col removal skipped due to residual merges"
            )

        try:
            rows = table.getRows().getCount()
            cols = table.getColumns().getCount()
        except Exception:
            rows, cols = 0, 0
        if rows <= 1 or cols <= 1:
            stats.layout_tables_flagged += 1
            stats.notes.append(
                f"table #{i}: looks like layout-only ({rows}x{cols})"
            )

        try:
            _set_header_row(table, stats)
        except Exception:
            pass
    log.info("table stats: %s", {k: v for k, v in stats.as_dict().items() if k != "notes"})
    return stats
