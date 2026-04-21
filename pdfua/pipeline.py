"""Pipeline orchestrator: input → ODT → PDF/UA + report."""

from __future__ import annotations

import logging
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from . import images as imagesmod
from . import normalizer, pdf_export, report, rules, spreadsheet, tables
from . import title_headings
from .uno_bridge import UnoBridge

log = logging.getLogger(__name__)

TEXT_EXTS = {".doc", ".docx", ".odt"}
SPREADSHEET_EXTS = {".xls", ".xlsx", ".ods"}


def detect_type(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in TEXT_EXTS:
        return "text"
    if ext in SPREADSHEET_EXTS:
        return "spreadsheet"
    raise ValueError(f"unsupported extension: {ext}")


def _safe_stem(s: str) -> str:
    s = re.sub(r"[^\w\- ()]+", "_", s, flags=re.UNICODE).strip()
    s = re.sub(r"\s+", " ", s)
    return s[:80] or "document"


def _stage(rep: report.PipelineReport, log_cb: Callable[[str], None] | None, name: str) -> None:
    rep.stages.append(name)
    log.info("STAGE: %s", name)
    if log_cb:
        log_cb(f"STAGE {name}")


@dataclass
class PipelineResult:
    odt_path: Path
    pdf_path: Path
    report_json_path: Path
    report_txt_path: Path
    report: report.PipelineReport


def run(input_path: Path, out_dir: Path, log_cb: Callable[[str], None] | None = None) -> PipelineResult:
    input_path = Path(input_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    kind = detect_type(input_path)

    rep = report.PipelineReport(
        input_path=str(input_path),
        input_type=kind,
        title="",
        title_source="",
        started_at=report.now(),
    )

    if log_cb:
        log_cb(f"STAGE bootstrap-soffice ({input_path.name}, type={kind})")

    with UnoBridge() as bridge:
        if log_cb:
            log_cb(f"soffice ready at {bridge.soffice_bin}")
        if kind == "text":
            result = _run_text(bridge, input_path, out_dir, rep, log_cb)
        else:
            result = _run_spreadsheet(bridge, input_path, out_dir, rep, log_cb)

    rep.finished_at = report.now()
    # Status
    errors = rep.rules.get("summary", {}).get("errors", 0) if rep.rules else 0
    warns = rep.rules.get("summary", {}).get("warnings", 0) if rep.rules else 0
    extra_risks = bool(rep.risks and (rep.tables.get("residual_merge_tables", 0) > 0
                                      or any("OCR" in r for r in rep.risks)))
    if errors == 0 and warns == 0 and not extra_risks:
        rep.status = "fully-fixed"
    elif errors == 0:
        rep.status = "fixed-with-warnings"
    else:
        rep.status = "best-effort-with-residual-risks"

    report.write_json(rep, result.report_json_path)
    report.write_text(rep, result.report_txt_path)
    if log_cb:
        log_cb(f"DONE status={rep.status}")
    return result


def _run_text(
    bridge: UnoBridge,
    src: Path,
    out_dir: Path,
    rep: report.PipelineReport,
    log_cb: Callable[[str], None] | None,
) -> PipelineResult:
    _stage(rep, log_cb, "intake")
    doc = bridge.load(src, hidden=True)
    doc_closed = False
    try:
        _stage(rep, log_cb, "title-detection")
        hstats = title_headings.ensure_title(doc, src)
        rep.title = hstats.title
        rep.title_source = hstats.title_source

        _stage(rep, log_cb, "normalize-text")
        n_stats = normalizer.normalize_document(doc)
        rep.normalizer = n_stats.as_dict()

        _stage(rep, log_cb, "restore-headings")
        hstats = title_headings.restore_headings(doc, hstats)
        rep.headings = hstats.as_dict()

        _stage(rep, log_cb, "repair-tables")
        t_stats = tables.repair_tables(doc)
        rep.tables = t_stats.as_dict()

        _stage(rep, log_cb, "process-images")
        work_dir = out_dir / "_work"
        work_dir.mkdir(exist_ok=True)
        i_stats = imagesmod.process_images(bridge, doc, work_dir)
        rep.images = i_stats.as_dict()

        _stage(rep, log_cb, "rule-check")
        rr = rules.check(doc)
        rep.rules = rr.as_dict()
        _collect_risks(rep, rr)

        _stage(rep, log_cb, "save-odt")
        stem = _safe_stem(rep.title or src.stem)
        odt_path = out_dir / f"{stem}.odt"
        bridge.save_as(doc, odt_path, "writer8")

        _stage(rep, log_cb, "export-pdfua")
        pdf_path = out_dir / f"{stem}.pdf"
        pdf_export.export_pdfua(bridge, doc, pdf_path, title=rep.title)
    finally:
        if not doc_closed and doc is not None:
            try:
                doc.close(True)
            except Exception:
                pass

    rep.outputs = {"odt": str(odt_path), "pdf": str(pdf_path)}
    json_path = out_dir / f"{stem}.report.json"
    txt_path = out_dir / f"{stem}.report.txt"

    # Cleanup work dir on success
    shutil.rmtree(out_dir / "_work", ignore_errors=True)

    return PipelineResult(odt_path, pdf_path, json_path, txt_path, rep)


def _run_spreadsheet(
    bridge: UnoBridge,
    src: Path,
    out_dir: Path,
    rep: report.PipelineReport,
    log_cb: Callable[[str], None] | None,
) -> PipelineResult:
    _stage(rep, log_cb, "intake-spreadsheet")

    # derive title from filename for first pass
    guessed_title = re.sub(r"[_\-]+", " ", src.stem).strip() or "Документ"
    rep.title = guessed_title
    rep.title_source = "filename"

    # 1. Calc → clean ODT
    intermediate_odt = out_dir / "_work" / "from_spreadsheet.odt"
    intermediate_odt.parent.mkdir(parents=True, exist_ok=True)

    _stage(rep, log_cb, "calc-repair-and-port-to-odt")
    sp_stats = spreadsheet.build_odt_from_spreadsheet(bridge, src, intermediate_odt, guessed_title)
    rep.spreadsheet = sp_stats.as_dict()

    # 2. Reopen that ODT and run the common text pipeline on it.
    doc = bridge.load(intermediate_odt, hidden=True)
    try:
        _stage(rep, log_cb, "title-detection")
        hstats = title_headings.ensure_title(doc, src)  # reuse file name semantics
        rep.title = hstats.title
        rep.title_source = hstats.title_source

        _stage(rep, log_cb, "normalize-text")
        n_stats = normalizer.normalize_document(doc)
        rep.normalizer = n_stats.as_dict()

        _stage(rep, log_cb, "restore-headings")
        hstats = title_headings.restore_headings(doc, hstats)
        rep.headings = hstats.as_dict()

        _stage(rep, log_cb, "repair-tables")
        t_stats = tables.repair_tables(doc)
        rep.tables = t_stats.as_dict()

        # Spreadsheet-derived ODT usually has no images, but cover the case.
        _stage(rep, log_cb, "process-images")
        work_dir = out_dir / "_work"
        i_stats = imagesmod.process_images(bridge, doc, work_dir)
        rep.images = i_stats.as_dict()

        _stage(rep, log_cb, "rule-check")
        rr = rules.check(doc)
        rep.rules = rr.as_dict()
        _collect_risks(rep, rr)

        _stage(rep, log_cb, "save-odt")
        stem = _safe_stem(rep.title or src.stem)
        odt_path = out_dir / f"{stem}.odt"
        bridge.save_as(doc, odt_path, "writer8")

        _stage(rep, log_cb, "export-pdfua")
        pdf_path = out_dir / f"{stem}.pdf"
        pdf_export.export_pdfua(bridge, doc, pdf_path, title=rep.title)
    finally:
        doc.close(True)

    rep.outputs = {
        "odt": str(odt_path),
        "pdf": str(pdf_path),
        "intermediate_odt": str(intermediate_odt),
    }
    json_path = out_dir / f"{stem}.report.json"
    txt_path = out_dir / f"{stem}.report.txt"

    return PipelineResult(odt_path, pdf_path, json_path, txt_path, rep)


def _collect_risks(rep: report.PipelineReport, rr: rules.RuleReport) -> None:
    for r in rr.results:
        if not r.passed:
            tag = r.severity.upper()
            rep.risks.append(f"[{tag}] {r.rule_id} {r.description}: {r.detail}")
    if rep.tables.get("residual_merge_tables", 0) > 0:
        rep.risks.append("Некоторые таблицы имеют остаточные объединения, требующие ручной доработки")
    low_conf_imgs = [
        r for r in rep.images.get("records", [])
        if r.get("ocr_confidence", 0) > 0 and not r.get("ocr_usable", False)
    ]
    if low_conf_imgs:
        rep.risks.append(
            f"{len(low_conf_imgs)} изображений: OCR с низкой уверенностью, alt text помечен как предположительный"
        )
