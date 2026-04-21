"""Flask web UI. Runs locally, no external calls.

Endpoints:
  GET  /                     index page
  POST /api/jobs             upload file, start job (runs in background thread)
  GET  /api/jobs/<id>        status + log snapshot
  GET  /api/jobs/<id>/<name> download a produced artifact
"""

from __future__ import annotations

import logging
import threading
import time
import uuid
from pathlib import Path
from typing import Any

from flask import Flask, abort, jsonify, request, send_from_directory

from . import pipeline

log = logging.getLogger(__name__)

WORK_ROOT = Path(__file__).resolve().parents[1] / "workdir" / "jobs"


class Job:
    def __init__(self, job_id: str, source: Path, out_dir: Path):
        self.id = job_id
        self.source = source
        self.out_dir = out_dir
        self.status = "pending"
        self.logs: list[str] = []
        self.started_at = time.time()
        self.finished_at: float | None = None
        self.result: pipeline.PipelineResult | None = None
        self.error: str | None = None
        self._lock = threading.Lock()

    def log(self, line: str) -> None:
        with self._lock:
            self.logs.append(f"[{time.strftime('%H:%M:%S')}] {line}")

    def to_public(self) -> dict:
        with self._lock:
            base: dict[str, Any] = {
                "id": self.id,
                "status": self.status,
                "source": self.source.name,
                "logs": list(self.logs),
                "elapsed": round((self.finished_at or time.time()) - self.started_at, 2),
                "error": self.error,
            }
            if self.result:
                r = self.result
                base["result"] = {
                    "title": r.report.title,
                    "title_source": r.report.title_source,
                    "status": r.report.status,
                    "risks": r.report.risks,
                    "downloads": {
                        "odt": r.odt_path.name,
                        "pdf": r.pdf_path.name,
                        "report_json": r.report_json_path.name,
                        "report_txt": r.report_txt_path.name,
                    },
                    "rules": r.report.rules,
                }
            return base


JOBS: dict[str, Job] = {}
JOBS_LOCK = threading.Lock()


def _run_job(job: Job) -> None:
    job.status = "running"
    job.log(f"start: {job.source.name}")
    try:
        def cb(m: str) -> None:
            job.log(m)
        result = pipeline.run(job.source, job.out_dir, log_cb=cb)
        job.result = result
        job.status = "done"
        job.log(f"finished: {result.report.status}")
    except Exception as e:
        log.exception("job %s failed", job.id)
        job.error = str(e)
        job.status = "error"
        job.log(f"ERROR: {e}")
    finally:
        job.finished_at = time.time()


def create_app() -> Flask:
    WORK_ROOT.mkdir(parents=True, exist_ok=True)
    app = Flask(__name__, static_folder="static", template_folder="templates")

    template_path = Path(__file__).resolve().parent / "templates" / "index.html"

    @app.get("/")
    def index():
        return template_path.read_text(encoding="utf-8")

    @app.post("/api/jobs")
    def create_job():
        if "file" not in request.files:
            abort(400, "missing file")
        f = request.files["file"]
        if not f.filename:
            abort(400, "empty filename")
        ext = Path(f.filename).suffix.lower()
        if ext not in {".doc", ".docx", ".odt", ".xls", ".xlsx", ".ods"}:
            abort(400, f"unsupported file type: {ext}")

        job_id = uuid.uuid4().hex[:12]
        job_dir = WORK_ROOT / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        src_path = job_dir / "input" / Path(f.filename).name
        src_path.parent.mkdir(parents=True, exist_ok=True)
        f.save(src_path)

        out_dir = job_dir / "output"
        job = Job(job_id, src_path, out_dir)
        with JOBS_LOCK:
            JOBS[job_id] = job
        t = threading.Thread(target=_run_job, args=(job,), daemon=True)
        t.start()
        return jsonify({"id": job_id})

    @app.get("/api/jobs/<job_id>")
    def get_job(job_id: str):
        job = JOBS.get(job_id)
        if not job:
            abort(404)
        return jsonify(job.to_public())

    @app.get("/api/jobs/<job_id>/download/<path:name>")
    def download(job_id: str, name: str):
        job = JOBS.get(job_id)
        if not job:
            abort(404)
        out_dir = job.out_dir
        # Only allow files from the output directory
        target = out_dir / name
        if not target.resolve().is_relative_to(out_dir.resolve()):
            abort(403)
        if not target.exists():
            abort(404)
        return send_from_directory(out_dir, name, as_attachment=True)

    return app


def main() -> None:  # pragma: no cover
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    app = create_app()
    app.run(host="127.0.0.1", port=8000, debug=False)


if __name__ == "__main__":  # pragma: no cover
    main()
