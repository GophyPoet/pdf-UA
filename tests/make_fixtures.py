"""Generate synthetic test documents that exercise every problem class
the pipeline is supposed to fix.

Uses the UNO bridge directly so the fixtures are saved via LibreOffice
itself and reflect realistic document internals.
"""

from __future__ import annotations

import logging
import os
import sys
from io import BytesIO
from pathlib import Path

# Make pdfua importable when running this script directly
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pdfua.uno_bridge import UnoBridge, path_to_url, make_prop  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")


FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _para(text, cursor, doc_text, style="Default Paragraph Style"):
    try:
        cursor.ParaStyleName = style
    except Exception:
        pass
    doc_text.insertString(cursor, text, False)
    doc_text.insertControlCharacter(cursor, 0, False)  # paragraph break


def make_docx_with_issues(bridge: UnoBridge, out: Path) -> None:
    """DOCX covering: multi-spaces, decorative tabs, underscore fields,
    pseudo-headings without styles, an image with text (for OCR)."""
    doc = bridge.new_writer()
    try:
        text = doc.getText()
        cursor = text.createTextCursorByRange(text.getEnd())

        # Pseudo-title: bold + large, no heading style
        try:
            cursor.ParaStyleName = "Default Paragraph Style"
            cursor.CharWeight = 150
            cursor.CharHeight = 18
            cursor.ParaAdjust = 3  # CENTER
        except Exception:
            pass
        text.insertString(cursor, "ИНСТРУКЦИЯ ПО ОФОРМЛЕНИЮ ДОКУМЕНТА", False)
        text.insertControlCharacter(cursor, 0, False)
        try:
            cursor.CharWeight = 100
            cursor.CharHeight = 11
            cursor.ParaAdjust = 0
        except Exception:
            pass

        # Body with multiple spaces and decorative tabs
        text.insertString(
            cursor,
            "Настоящая   инструкция\tописывает   порядок\tоформления    документа.",
            False,
        )
        text.insertControlCharacter(cursor, 0, False)

        # Pseudo section heading
        try:
            cursor.CharWeight = 150
            cursor.CharHeight = 14
        except Exception:
            pass
        text.insertString(cursor, "1. Общие положения", False)
        text.insertControlCharacter(cursor, 0, False)
        try:
            cursor.CharWeight = 100
            cursor.CharHeight = 11
        except Exception:
            pass

        text.insertString(
            cursor,
            "Документ оформляется в строгом соответствии с требованиями. "
            "Внутри абзацев  не должно быть  двойных пробелов.",
            False,
        )
        text.insertControlCharacter(cursor, 0, False)

        # Underscore fill line
        text.insertString(cursor, "ФИО: ______________________________", False)
        text.insertControlCharacter(cursor, 0, False)
        text.insertString(cursor, "__________________________________________", False)
        text.insertControlCharacter(cursor, 0, False)

        # Table with a merged cell and an empty cell
        table = doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(3, 3)
        text.insertTextContent(cursor, table, False)
        table.getCellByName("A1").setString("Столбец 1")
        table.getCellByName("B1").setString("Столбец 2")
        table.getCellByName("C1").setString("Столбец 3")
        table.getCellByName("A2").setString("Данные 1")
        # B2 left empty on purpose
        table.getCellByName("C2").setString("Данные 3")
        table.getCellByName("A3").setString("Данные 4")
        table.getCellByName("B3").setString("Данные 5")
        table.getCellByName("C3").setString("Данные 6")
        # Merge A2..B2 to create a merged cell
        try:
            curs = table.createCursorByCellName("A2")
            curs.goRight(1, True)
            curs.mergeRange()
        except Exception:
            pass

        # Paragraph after table
        text.insertControlCharacter(text.createTextCursorByRange(text.getEnd()), 0, False)

        # Insert an image with text via a PNG we build with PIL
        from PIL import Image, ImageDraw, ImageFont
        img = Image.new("RGB", (600, 200), "white")
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype(
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 28
            )
        except Exception:
            font = ImageFont.load_default()
        draw.text((10, 20), "СКРИНШОТ ДОКУМЕНТА", fill="black", font=font)
        draw.text((10, 80), "Очень важная информация.", fill="black", font=font)
        draw.text((10, 140), "Читайте ниже в отчёте.", fill="black", font=font)
        img_path = FIXTURES / "tmp_image.png"
        img.save(img_path, "PNG")

        # Insert image using UNO GraphicProvider
        from com.sun.star.awt import Size  # type: ignore
        smgr = bridge.ctx.ServiceManager
        provider = smgr.createInstanceWithContext(
            "com.sun.star.graphic.GraphicProvider", bridge.ctx
        )
        url = path_to_url(img_path)
        gr = provider.queryGraphic((make_prop("URL", url),))

        shape = doc.createInstance("com.sun.star.text.TextGraphicObject")
        shape.Graphic = gr
        sz = Size()
        sz.Width = 10000
        sz.Height = 3500
        shape.Size = sz
        shape.AnchorType = 1  # AS_CHARACTER
        cursor_end = text.createTextCursorByRange(text.getEnd())
        text.insertTextContent(cursor_end, shape, False)
        text.insertControlCharacter(text.createTextCursorByRange(text.getEnd()), 0, False)

        text.insertString(
            text.createTextCursorByRange(text.getEnd()),
            "На этом инструкция завершается.",
            False,
        )

        # Save DOCX
        bridge.save_as(doc, out, "MS Word 2007 XML")
    finally:
        doc.close(True)


def make_xlsx_with_issues(bridge: UnoBridge, out: Path) -> None:
    """XLSX with multiple sheets, merged cells, empty columns."""
    doc = bridge.desktop.loadComponentFromURL(
        "private:factory/scalc", "_blank", 0, (make_prop("Hidden", True),)
    )
    try:
        sheets = doc.getSheets()
        # Sheet 1: realistic data + one merged header + empty column C
        s1 = sheets.getByIndex(0)
        s1.setName("Финансы Q1")
        s1.getCellByPosition(0, 0).setString("Отчёт за первый квартал")
        # Merge A1:D1 as big title
        try:
            rng = s1.getCellRangeByPosition(0, 0, 3, 0)
            rng.merge(True)
        except Exception:
            pass
        headers = ["Месяц", "Доход", "", "Расход"]
        for i, h in enumerate(headers):
            s1.getCellByPosition(i, 2).setString(h)
        for r, (m, d, e) in enumerate([
            ("Январь", 100, 60),
            ("Февраль", 120, 70),
            ("Март", 150, 80),
        ]):
            s1.getCellByPosition(0, 3 + r).setString(m)
            s1.getCellByPosition(1, 3 + r).setValue(d)
            # skip col 2 (empty column)
            s1.getCellByPosition(3, 3 + r).setValue(e)

        # Sheet 2: small sheet
        if sheets.getCount() < 2:
            sheets.insertNewByName("Примечания", 1)
        else:
            s2 = sheets.getByIndex(1)
            s2.setName("Примечания")
        s2 = sheets.getByName("Примечания")
        s2.getCellByPosition(0, 0).setString("Примечания к отчёту")
        s2.getCellByPosition(0, 2).setString("Все суммы в тысячах рублей.")

        bridge.save_as(doc, out, "Calc MS Excel 2007 XML")
    finally:
        doc.close(True)


def main() -> None:
    FIXTURES.mkdir(parents=True, exist_ok=True)
    with UnoBridge() as b:
        make_docx_with_issues(b, FIXTURES / "sample_doc.docx")
        make_xlsx_with_issues(b, FIXTURES / "sample_sheet.xlsx")
    # Clean tmp image
    tmp = FIXTURES / "tmp_image.png"
    if tmp.exists():
        tmp.unlink()
    print("fixtures:")
    for p in sorted(FIXTURES.glob("*")):
        print(" ", p)


if __name__ == "__main__":
    main()
