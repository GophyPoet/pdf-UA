"""End-to-end integration tests.

Requires LibreOffice + Tesseract. Generates fixtures on the fly if
missing and runs the full pipeline for both text and spreadsheet paths.
"""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pdfua import pipeline

ROOT = Path(__file__).resolve().parent
FIXTURES = ROOT / "fixtures"
OUT = ROOT / "_e2e_out"


def _ensure_fixtures():
    if not (FIXTURES / "sample_doc.docx").exists() or not (FIXTURES / "sample_sheet.xlsx").exists():
        from tests import make_fixtures  # type: ignore
        make_fixtures.main()


def _fresh_outdir(name: str) -> Path:
    d = OUT / name
    if d.exists():
        shutil.rmtree(d)
    d.mkdir(parents=True)
    return d


def test_docx_full_pipeline():
    _ensure_fixtures()
    # remove any stale lock
    for p in FIXTURES.glob(".~lock.*"):
        p.unlink()
    out_dir = _fresh_outdir("docx")
    result = pipeline.run(FIXTURES / "sample_doc.docx", out_dir)
    r = result.report

    # Outputs present
    assert result.odt_path.exists(), "ODT missing"
    assert result.pdf_path.exists(), "PDF missing"
    assert result.report_json_path.exists(), "JSON report missing"
    assert result.report_txt_path.exists(), "TXT report missing"

    # PDF is 1.7 and has PDF/UA marker
    data = result.pdf_path.read_bytes()
    assert data.startswith(b"%PDF-1.7"), "PDF is not 1.7"
    assert b"pdfuaid:part" in data, "PDF/UA marker missing"

    # Title carried through
    assert r.title, "title not set"
    # Normalizer saw multi-spaces and tabs in the fixture
    assert r.normalizer.get("multi_spaces_collapsed", 0) > 0
    assert r.normalizer.get("tabs_removed", 0) > 0
    # Image was found and described
    assert r.images.get("total", 0) == 1
    assert r.images.get("with_usable_ocr", 0) == 1
    assert r.images.get("newly_described", 0) == 1
    # Table was processed
    assert r.tables.get("total", 0) == 1
    assert r.tables.get("header_rows_set", 0) == 1
    # Rules: no ERROR-level failures
    summary = r.rules.get("summary", {})
    assert summary.get("errors", 0) == 0, f"rule errors: {r.rules}"


def test_xlsx_full_pipeline():
    _ensure_fixtures()
    for p in FIXTURES.glob(".~lock.*"):
        p.unlink()
    out_dir = _fresh_outdir("xlsx")
    result = pipeline.run(FIXTURES / "sample_sheet.xlsx", out_dir)
    r = result.report
    assert result.odt_path.exists()
    assert result.pdf_path.exists()

    # 2 sheets → at least 2 Writer tables
    sp = r.spreadsheet
    assert sp.get("total_sheets", 0) == 2
    assert sp.get("total_merged_regions", 0) >= 1, "merge should have been handled"

    # Rules pass
    summary = r.rules.get("summary", {})
    assert summary.get("errors", 0) == 0


def _run(name):
    import traceback
    fn = globals()[name]
    try:
        fn()
        print("  OK", name)
        return True
    except AssertionError as e:
        print("FAIL", name, "-", e)
        return False
    except Exception:
        print("FAIL", name)
        traceback.print_exc()
        return False


if __name__ == "__main__":
    ok = 0
    fail = 0
    for name in ("test_docx_full_pipeline", "test_xlsx_full_pipeline"):
        if _run(name):
            ok += 1
        else:
            fail += 1
    print(f"{ok} passed, {fail} failed")
    sys.exit(0 if fail == 0 else 1)
