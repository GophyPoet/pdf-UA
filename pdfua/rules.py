"""Accessibility rule engine.

A straightforward rule set that inspects the final ODT document and
reports remaining risks. We try to keep it honest: rules return
`severity` and `passed` booleans, not a single green/red flag.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


SEVERITY_ERROR = "error"
SEVERITY_WARNING = "warning"
SEVERITY_INFO = "info"


@dataclass
class RuleResult:
    rule_id: str
    description: str
    passed: bool
    severity: str
    detail: str = ""

    def as_dict(self) -> dict:
        return dict(self.__dict__)


@dataclass
class RuleReport:
    results: list[RuleResult] = field(default_factory=list)
    ready_for_pdfua: bool = False

    def add(self, r: RuleResult) -> None:
        self.results.append(r)

    def as_dict(self) -> dict:
        return {
            "ready_for_pdfua": self.ready_for_pdfua,
            "summary": {
                "errors": sum(1 for r in self.results if not r.passed and r.severity == SEVERITY_ERROR),
                "warnings": sum(1 for r in self.results if not r.passed and r.severity == SEVERITY_WARNING),
                "passed": sum(1 for r in self.results if r.passed),
                "total": len(self.results),
            },
            "results": [r.as_dict() for r in self.results],
        }


def _iter_paragraphs(doc):
    enum = doc.getText().createEnumeration()
    while enum.hasMoreElements():
        el = enum.nextElement()
        if el.supportsService("com.sun.star.text.Paragraph"):
            yield el
        elif el.supportsService("com.sun.star.text.TextTable"):
            for n in el.getCellNames():
                sub = el.getCellByName(n).createEnumeration()
                while sub.hasMoreElements():
                    pel = sub.nextElement()
                    if pel.supportsService("com.sun.star.text.Paragraph"):
                        yield pel


def check(doc) -> RuleReport:
    report = RuleReport()

    # R1: Title
    title = ""
    try:
        title = (doc.getDocumentProperties().Title or "").strip()
    except Exception:
        pass
    report.add(RuleResult(
        "R001", "Document has a Title property",
        passed=bool(title), severity=SEVERITY_ERROR,
        detail=f"Title={title!r}",
    ))

    # R2: At least one heading
    heading_count = 0
    for p in _iter_paragraphs(doc):
        try:
            style = p.ParaStyleName or ""
        except Exception:
            style = ""
        if style.startswith("Heading") or style == "Title":
            heading_count += 1
    report.add(RuleResult(
        "R002", "Document has at least one heading",
        passed=heading_count > 0, severity=SEVERITY_ERROR,
        detail=f"headings={heading_count}",
    ))

    # R3: No runs of multiple spaces inside paragraphs
    multi_space_hits = 0
    for p in _iter_paragraphs(doc):
        if re.search(r"  +", p.getString() or ""):
            multi_space_hits += 1
    report.add(RuleResult(
        "R003", "No multiple-space runs",
        passed=multi_space_hits == 0, severity=SEVERITY_WARNING,
        detail=f"paragraphs_with_multi_spaces={multi_space_hits}",
    ))

    # R4: No decorative tabs inside paragraphs
    tab_hits = 0
    for p in _iter_paragraphs(doc):
        if "\t" in (p.getString() or ""):
            tab_hits += 1
    report.add(RuleResult(
        "R004", "No decorative tabs inside paragraphs",
        passed=tab_hits == 0, severity=SEVERITY_WARNING,
        detail=f"paragraphs_with_tabs={tab_hits}",
    ))

    # R5: No long underscore chains
    underscore_hits = 0
    for p in _iter_paragraphs(doc):
        if re.search(r"_{3,}", p.getString() or ""):
            underscore_hits += 1
    report.add(RuleResult(
        "R005", "No 3+ underscore fill lines",
        passed=underscore_hits == 0, severity=SEVERITY_WARNING,
        detail=f"paragraphs_with_underscore_runs={underscore_hits}",
    ))

    # R6/R7: Tables
    tables_with_merges = 0
    tables_with_empty_cells = 0
    empty_rows_found = 0
    tables = doc.getTextTables()
    for i in range(tables.getCount()):
        t = tables.getByIndex(i)
        names = t.getCellNames()
        if any("." in n for n in names):
            tables_with_merges += 1
        for cn in names:
            try:
                if not t.getCellByName(cn).getString().strip():
                    tables_with_empty_cells += 1
                    break
            except Exception:
                continue
        try:
            rows = t.getRows().getCount()
            cols = t.getColumns().getCount()
        except Exception:
            continue
        letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        for r in range(rows):
            empty = True
            for c in range(cols):
                cn = f"{letters[c % 26]}{r + 1}"
                try:
                    if t.getCellByName(cn).getString().strip():
                        empty = False
                        break
                except Exception:
                    continue
            if empty:
                empty_rows_found += 1

    report.add(RuleResult(
        "R006", "Tables have no merged cells",
        passed=tables_with_merges == 0, severity=SEVERITY_ERROR,
        detail=f"tables_with_merges={tables_with_merges}",
    ))
    report.add(RuleResult(
        "R007", "Tables have no empty cells",
        passed=tables_with_empty_cells == 0, severity=SEVERITY_WARNING,
        detail=f"tables_with_empty_cells={tables_with_empty_cells}",
    ))
    report.add(RuleResult(
        "R008", "Tables have no empty rows",
        passed=empty_rows_found == 0, severity=SEVERITY_WARNING,
        detail=f"empty_rows={empty_rows_found}",
    ))

    # R9: Every image has alt text or is marked decorative
    try:
        graphics = doc.getGraphicObjects()
    except Exception:
        graphics = None
    missing_alt = 0
    image_total = 0
    if graphics is not None:
        for i in range(graphics.getCount()):
            image_total += 1
            shape = graphics.getByIndex(i)
            alt = ""
            try:
                alt = (shape.Description or "").strip()
            except Exception:
                pass
            if not alt:
                try:
                    alt = (shape.Title or "").strip()
                except Exception:
                    pass
            if not alt:
                missing_alt += 1
    report.add(RuleResult(
        "R009", "All images have alt text or are marked decorative",
        passed=missing_alt == 0, severity=SEVERITY_ERROR,
        detail=f"images={image_total} missing_alt={missing_alt}",
    ))

    # Determine readiness: no ERROR-level failures
    errors = [r for r in report.results if not r.passed and r.severity == SEVERITY_ERROR]
    report.ready_for_pdfua = len(errors) == 0
    return report
