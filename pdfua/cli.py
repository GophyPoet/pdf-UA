"""Command-line entry point.

Usage:
    python -m pdfua.cli convert input.docx output_dir/
    python -m pdfua.cli serve [--host 0.0.0.0 --port 8000]
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from . import pipeline


def _cli_log(msg: str) -> None:
    print(msg, flush=True)


def _cmd_convert(args: argparse.Namespace) -> int:
    src = Path(args.input)
    if not src.exists():
        print(f"error: input not found: {src}", file=sys.stderr)
        return 2
    out_dir = Path(args.output)
    result = pipeline.run(src, out_dir, log_cb=_cli_log)
    print()
    print(f"status:     {result.report.status}")
    print(f"title:      {result.report.title}  (source: {result.report.title_source})")
    print(f"odt:        {result.odt_path}")
    print(f"pdf:        {result.pdf_path}")
    print(f"report.json:{result.report_json_path}")
    print(f"report.txt: {result.report_txt_path}")
    if result.report.risks:
        print()
        print("Остаточные риски:")
        for r in result.report.risks:
            print(f"  - {r}")
    return 0


def _cmd_serve(args: argparse.Namespace) -> int:
    from .server import create_app  # lazy import
    app = create_app()
    if args.open:
        import threading
        import time
        import webbrowser

        url = f"http://{args.host if args.host != '0.0.0.0' else '127.0.0.1'}:{args.port}/"

        def _open() -> None:
            time.sleep(1.0)  # give Flask a moment to bind
            try:
                webbrowser.open(url)
            except Exception:
                pass

        threading.Thread(target=_open, daemon=True).start()
        print(f"opening browser at {url}")
    app.run(host=args.host, port=args.port, debug=False)
    return 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="pdfua", description="PDF/UA автоматизация через LibreOffice")
    sub = p.add_subparsers(dest="cmd", required=True)

    pc = sub.add_parser("convert", help="конвертировать один файл")
    pc.add_argument("input", help="входной .doc/.docx/.odt/.xls/.xlsx/.ods")
    pc.add_argument("output", help="выходная папка")
    pc.add_argument("-v", "--verbose", action="store_true")
    pc.set_defaults(func=_cmd_convert)

    ps = sub.add_parser("serve", help="запустить web-UI")
    ps.add_argument("--host", default="127.0.0.1")
    ps.add_argument("--port", default=8000, type=int)
    ps.add_argument("--open", action="store_true",
                    help="автоматически открыть UI в браузере")
    ps.set_defaults(func=_cmd_serve)

    args = p.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if getattr(args, "verbose", False) else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
