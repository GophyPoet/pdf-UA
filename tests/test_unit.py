"""Unit-level tests — no soffice needed."""

from __future__ import annotations

import sys
from io import BytesIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PIL import Image, ImageDraw, ImageFont

from pdfua import alt_text as altmod
from pdfua.ocr import run_ocr


def test_alt_decorative_for_tiny_image():
    from pdfua.ocr import OcrResult
    decision = altmod.decide(OcrResult("", 0.0, "rus+eng", 0, False), 30, 30)
    assert decision.decorative is True
    assert decision.text_equivalent == ""


def test_alt_rule_line_marked_decorative():
    from pdfua.ocr import OcrResult
    decision = altmod.decide(OcrResult("", 0.0, "rus+eng", 0, False), 500, 4)
    assert decision.decorative is True


def test_alt_from_readable_ocr():
    from pdfua.ocr import OcrResult
    res = OcrResult("Пример заголовка. Тело текста.", 0.82, "rus+eng", 5, True)
    decision = altmod.decide(res, 800, 400)
    assert decision.decorative is False
    assert "Пример заголовка" in decision.alt_text


def test_alt_low_confidence_marks_uncertainty():
    from pdfua.ocr import OcrResult
    res = OcrResult("неразборчивый текст", 0.35, "rus+eng", 3, False)
    decision = altmod.decide(res, 800, 400)
    assert decision.decorative is False
    assert "предположительно" in decision.alt_text
    assert decision.text_equivalent == ""


def test_ocr_on_synthetic_text():
    img = Image.new("RGB", (900, 200), "white")
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 36
        )
    except Exception:
        font = ImageFont.load_default()
    draw.text((20, 60), "HELLO OCR TEST", fill="black", font=font)
    buf = BytesIO()
    img.save(buf, "PNG")
    res = run_ocr(buf.getvalue())
    assert res.usable, f"OCR failed: conf={res.confidence} text={res.text!r}"
    assert "HELLO" in res.text.upper()


def test_normalizer_regex_collapses_multi_spaces():
    from pdfua.normalizer import MULTI_SPACE_RE, TAB_RE, UNDERSCORE_LINE_RE
    assert MULTI_SPACE_RE.sub(" ", "a   b") == "a b"
    assert TAB_RE.sub(" ", "a\t\tb") == "a b"
    assert UNDERSCORE_LINE_RE.search("___") is not None
    assert UNDERSCORE_LINE_RE.search("__") is None


if __name__ == "__main__":  # pragma: no cover
    import traceback
    ok = 0
    fail = 0
    for name, fn in list(globals().items()):
        if not name.startswith("test_") or not callable(fn):
            continue
        try:
            fn()
            print("  OK", name)
            ok += 1
        except Exception:
            print("FAIL", name)
            traceback.print_exc()
            fail += 1
    print(f"{ok} passed, {fail} failed")
    sys.exit(0 if fail == 0 else 1)
