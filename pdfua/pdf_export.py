"""PDF/UA export.

Uses Writer's 'writer_pdf_Export' filter with an explicit FilterData
map. The parameters are fixed by requirement:

  - all pages
  - JPEG image compression at 90% quality
  - reduce image resolution to 150 DPI
  - PDF version 1.7
  - PDF/UA enabled
  - tagged / structured PDF enabled
"""

from __future__ import annotations

import logging
import os
import shutil
import tempfile
import threading
import time
from pathlib import Path

import uno  # type: ignore
from com.sun.star.beans import PropertyValue  # type: ignore

from .uno_bridge import UnoBridge, path_to_url, make_prop

log = logging.getLogger(__name__)


def _is_ascii(s: str) -> bool:
    try:
        s.encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def _heartbeat(stop: threading.Event, label: str, interval: float = 15.0) -> None:
    t0 = time.time()
    while not stop.wait(interval):
        log.info("%s still running after %.0fs", label, time.time() - t0)


def export_pdfua(bridge: UnoBridge, doc, dest_pdf: Path, title: str | None = None) -> Path:
    # 1. Mirror title into document metadata so the PDF carries it.
    if title:
        try:
            doc.getDocumentProperties().Title = title
        except Exception:
            pass

    # 2. Build FilterData
    filter_data = [
        make_prop("SelectPdfVersion", 17),        # PDF 1.7
        make_prop("PDFUACompliance", True),
        make_prop("UseTaggedPDF", True),
        make_prop("ExportBookmarks", True),
        make_prop("ExportNotes", False),
        make_prop("UseLosslessCompression", False),
        make_prop("Quality", 90),                 # JPEG quality
        make_prop("ReduceImageResolution", True),
        make_prop("MaxImageResolution", 150),
        make_prop("ExportFormFields", True),
        make_prop("IsSkipEmptyPages", False),
        make_prop("ExportLinksRelativeFsys", False),
    ]
    # "Selection" / "PageRange" intentionally omitted -> export all pages.

    fd_any = uno.Any("[]com.sun.star.beans.PropertyValue", tuple(filter_data))
    store_props = (
        make_prop("FilterName", "writer_pdf_Export"),
        make_prop("Overwrite", True),
        make_prop("FilterData", fd_any),
    )

    # 3. Export to an ASCII-only temp path first, then move into place.
    #
    # On Windows, writing a PDF directly to a path that contains non-ASCII
    # characters (Cyrillic, in particular) or to a OneDrive-synced folder
    # like %USERPROFILE%\Desktop can make storeToURL hang for minutes.
    # Exporting to a stable local %TEMP% path (ASCII) and then renaming
    # makes the operation reliable and gives us a single small rename at
    # the end instead of many small file-sync round-trips during encoding.
    dest_pdf = Path(dest_pdf)
    needs_relay = (
        os.name == "nt"
        and not _is_ascii(str(dest_pdf))
    )

    if needs_relay:
        fd, tmp_name = tempfile.mkstemp(prefix="pdfua_export_", suffix=".pdf")
        os.close(fd)
        tmp_path = Path(tmp_name)
        log.info("exporting PDF/UA via ASCII relay: %s -> %s", tmp_path, dest_pdf)
    else:
        tmp_path = dest_pdf
        log.info("exporting PDF/UA to %s", dest_pdf)

    url = path_to_url(tmp_path)

    # 4. Run storeToURL with a heartbeat so we can see it's still alive.
    stop = threading.Event()
    hb = threading.Thread(
        target=_heartbeat, args=(stop, "storeToURL (PDF/UA export)"), daemon=True
    )
    hb.start()
    try:
        doc.storeToURL(url, store_props)
    finally:
        stop.set()
        hb.join(timeout=1.0)

    # 5. Relay temp file to final location if we used one.
    if needs_relay:
        dest_pdf.parent.mkdir(parents=True, exist_ok=True)
        if dest_pdf.exists():
            dest_pdf.unlink()
        shutil.move(str(tmp_path), str(dest_pdf))

    return dest_pdf
