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
from pathlib import Path

import uno  # type: ignore
from com.sun.star.beans import PropertyValue  # type: ignore

from .uno_bridge import UnoBridge, path_to_url, make_prop

log = logging.getLogger(__name__)


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
    url = path_to_url(dest_pdf)
    log.info("exporting PDF/UA to %s", dest_pdf)
    doc.storeToURL(url, store_props)
    return dest_pdf
