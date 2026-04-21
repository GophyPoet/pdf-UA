"""Report builder.

Produces a JSON file and a human-readable text report side by side.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


@dataclass
class PipelineReport:
    input_path: str
    input_type: str
    title: str
    title_source: str
    stages: list[str] = field(default_factory=list)
    normalizer: dict = field(default_factory=dict)
    headings: dict = field(default_factory=dict)
    images: dict = field(default_factory=dict)
    tables: dict = field(default_factory=dict)
    spreadsheet: dict = field(default_factory=dict)
    rules: dict = field(default_factory=dict)
    outputs: dict = field(default_factory=dict)
    risks: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    status: str = "pending"
    started_at: str = ""
    finished_at: str = ""

    def as_dict(self) -> dict:
        return asdict(self)


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def write_json(report: PipelineReport, path: Path) -> None:
    path.write_text(
        json.dumps(report.as_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _section(title: str) -> str:
    line = "=" * max(3, 40 - len(title))
    return f"\n== {title} {line}\n"


def _bool(v: Any) -> str:
    return "да" if v else "нет"


def write_text(report: PipelineReport, path: Path) -> None:
    r = report
    lines: list[str] = []
    lines.append("Отчёт подготовки документа к PDF/UA")
    lines.append(f"Время: {r.started_at} → {r.finished_at}")
    lines.append(f"Исходный файл: {r.input_path}")
    lines.append(f"Тип: {r.input_type}")
    lines.append(f"Заглавие: {r.title}  (источник: {r.title_source})")
    lines.append(f"Статус: {r.status}")

    lines.append(_section("Этапы"))
    for s in r.stages:
        lines.append(f" • {s}")

    if r.normalizer:
        lines.append(_section("Нормализация текста"))
        for k, v in r.normalizer.items():
            lines.append(f" • {k}: {v}")

    if r.headings:
        lines.append(_section("Заголовки и заглавие"))
        for k, v in r.headings.items():
            lines.append(f" • {k}: {v}")

    if r.tables:
        lines.append(_section("Таблицы (Writer)"))
        for k, v in r.tables.items():
            if k == "notes":
                continue
            lines.append(f" • {k}: {v}")
        for n in r.tables.get("notes", []):
            lines.append(f"   ⚠ {n}")

    if r.spreadsheet:
        lines.append(_section("Таблицы (Calc → Writer)"))
        for k, v in r.spreadsheet.items():
            if k == "sheets":
                continue
            lines.append(f" • {k}: {v}")
        for s in r.spreadsheet.get("sheets", []):
            lines.append(
                f"   ○ Лист '{s['name']}': rows={s.get('rows',0)} cols={s.get('cols',0)}, "
                f"header_filled={s.get('header_cells_filled',0)}, "
                f"truncated={s.get('cells_truncated',0)}, "
                f"empty_filled={s.get('cells_filled',0)}, "
                f"empty_rows_skipped={s.get('empty_rows_skipped',0)}, "
                f"empty_cols_skipped={s.get('empty_cols_skipped',0)}"
                + (", SKIPPED" if s.get('skipped') else "")
            )

    if r.images:
        lines.append(_section("Изображения / OCR / alt text"))
        for k, v in r.images.items():
            if k == "records":
                continue
            lines.append(f" • {k}: {v}")
        for rec in r.images.get("records", []):
            lines.append(
                f"   ○ #{rec['index']} {rec['name']} "
                f"({rec['width_px']}x{rec['height_px']}): "
                f"alt={rec['assigned_alt']!r}; "
                f"decorative={_bool(rec['decorative'])}; "
                f"ocr_conf={rec['ocr_confidence']}; "
                f"ocr_usable={_bool(rec['ocr_usable'])}; "
                f"text_equivalent={_bool(rec['text_equivalent_injected'])}; "
                f"reason={rec['reasoning']}"
            )
            if rec["ocr_text"]:
                first = rec["ocr_text"].splitlines()[0]
                lines.append(f"     OCR: {first!r}")

    if r.rules:
        lines.append(_section("Проверка доступности"))
        lines.append(f"PDF/UA ready: {_bool(r.rules.get('ready_for_pdfua'))}")
        summ = r.rules.get("summary", {})
        lines.append(
            f"Всего правил: {summ.get('total',0)}, "
            f"прошло: {summ.get('passed',0)}, "
            f"ошибки: {summ.get('errors',0)}, "
            f"предупреждения: {summ.get('warnings',0)}"
        )
        for rr in r.rules.get("results", []):
            mark = "OK" if rr["passed"] else rr["severity"].upper()
            lines.append(f"   [{mark}] {rr['rule_id']}: {rr['description']} — {rr['detail']}")

    if r.risks:
        lines.append(_section("Остаточные риски"))
        for risk in r.risks:
            lines.append(f" ! {risk}")

    if r.warnings:
        lines.append(_section("Предупреждения пайплайна"))
        for w in r.warnings:
            lines.append(f" ~ {w}")

    if r.outputs:
        lines.append(_section("Результаты"))
        for k, v in r.outputs.items():
            lines.append(f" • {k}: {v}")

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
