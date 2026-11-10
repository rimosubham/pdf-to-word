"""PDF -> Word conversion.

Two engines:
  * layout  - pdf2docx, for PDFs that carry a real text layer. Keeps fonts,
              tables, images, columns and page geometry.
  * ocr     - for scanned / image-only pages: Apple Vision on macOS, Tesseract
              elsewhere (e.g. on a Linux server). Text is
              recognised and rebuilt as paragraphs; the layout is simplified.

A document is routed to OCR only when most of its pages have no usable text
layer. In OCR mode, pages that *do* have a text layer still use it directly.
"""
from __future__ import annotations

import logging
import os
import statistics
import sys
import warnings

warnings.filterwarnings("ignore")

import pymupdf  # noqa: E402
from docx import Document  # noqa: E402
from docx.shared import Pt  # noqa: E402
from pdf2docx import Converter  # noqa: E402

logging.getLogger("pdf2docx").setLevel(logging.ERROR)
logging.getLogger().setLevel(logging.ERROR)

MIN_CHARS_FOR_TEXT_PAGE = 25
OCR_DPI = int(os.environ.get("OCR_DPI", 300))


class ConversionError(Exception):
    """Raised with a message that is safe to show to the user."""


def inspect_pdf(pdf_path: str, password: str | None = None) -> dict:
    """Page count and how many pages lack a text layer."""
    try:
        doc = pymupdf.open(pdf_path)
    except Exception as exc:  # corrupt / not a PDF
        raise ConversionError("This file could not be opened as a PDF.") from exc
    try:
        if doc.needs_pass:
            if not password or not doc.authenticate(password):
                raise ConversionError(
                    "This PDF is password-protected. Enter the password and try again."
                    if not password
                    else "That password did not unlock the PDF."
                )
        if doc.page_count == 0:
            raise ConversionError("This PDF has no pages.")
        scanned = []
        for page in doc:
            visible = sum(len(s["text"].strip())
                          for b in page.get_text("dict")["blocks"]
                          for l in b.get("lines", [])
                          for s in l["spans"] if s.get("alpha", 255) > 0)
            if visible < MIN_CHARS_FOR_TEXT_PAGE:
                # Blank pages are not "scanned": only count pages that show an image.
                if page.get_images(full=False) or page.get_drawings():
                    scanned.append(page.number)
        return {"pages": doc.page_count, "scanned_pages": scanned}
    finally:
        doc.close()


def convert(pdf_path: str, docx_path: str, password: str | None = None,
            progress=None) -> dict:
    """Convert, verify, repair. Returns engine details plus the accuracy report."""
    import accuracy
    import repair

    info = inspect_pdf(pdf_path, password)
    pages, scanned = info["pages"], set(info["scanned_pages"])
    use_ocr = len(scanned) > pages / 2
    repaired = 0

    if use_ocr:
        ocr_cache: dict[int, list[dict]] = {}     # recognise each page once
        _convert_ocr(pdf_path, docx_path, password, scanned, progress, ocr_cache)
        lines = accuracy.pdf_lines(pdf_path, password, scanned, ocr_cache)
    else:
        lines = accuracy.pdf_lines(pdf_path, password, scanned)

        # Scanned pages inside a text PDF stay as pictures; don't paste OCR text in.
        fixable = [l for l in lines if l["page"] not in scanned]
        fixed: dict[str, int] = {}

        def kept(path):
            """Repair the attempt, then report the share of words it retains."""
            fixed[path] = repair.repair(path, fixable)
            rep = accuracy.compare(pdf_path, path, password, scanned, lines)
            return (rep["recall"] if rep["accuracy"] is not None else 100.0), fixed[path]

        if progress:
            progress(0, pages, "Rebuilding layout and checking every word")
        best = _convert_layout(pdf_path, docx_path, password, kept,
                               tolerable_repairs=max(2, len(fixable) // 50))
        repaired = fixed.get(best, 0)

    report = accuracy.compare(pdf_path, docx_path, password, scanned, lines)
    return {"engine": "ocr" if use_ocr else "layout", "pages": pages,
            "scanned_pages": len(scanned), "repaired_lines": repaired, **report}


# Settings tried in order until the Word file keeps (nearly) every word.
LAYOUT_ATTEMPTS = [
    {},
    {"parse_lattice_table": False},
    {"parse_lattice_table": False, "parse_stream_table": False},
]
GOOD_ENOUGH = 99.0


def _convert_layout(pdf_path: str, docx_path: str, password: str | None,
                    score=None, tolerable_repairs: int = 2) -> str:
    """Run pdf2docx. `score(docx)` repairs the attempt and returns
    (percent of words kept, lines it had to re-insert). When an attempt loses
    text or needs many re-insertions (a sign a table was mangled), try other
    table settings and keep the cleanest result."""
    done, last_error = [], None
    for n, settings in enumerate(LAYOUT_ATTEMPTS):
        attempt = f"{docx_path}.try{n}"
        cv = Converter(pdf_path, password=password) if password else Converter(pdf_path)
        try:
            cv.convert(attempt, **settings)
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            continue
        finally:
            cv.close()
        kept, fixed = score(attempt) if score else (100.0, 0)
        done.append((kept < GOOD_ENOUGH, fixed, -kept, attempt))
        if kept >= GOOD_ENOUGH and fixed <= tolerable_repairs:
            break
    if not done:
        raise ConversionError(f"The converter failed on this PDF ({last_error}).") from last_error
    best = min(done)[3]
    for *_k, path in done:
        if path != best:
            os.remove(path)
    os.replace(best, docx_path)
    return best


# ---------------------------------------------------------------- OCR engine

def _ocr_engine() -> str:
    """'vision' (Apple, macOS only) or 'tesseract' (anywhere). OCR_ENGINE overrides."""
    forced = os.environ.get("OCR_ENGINE")
    if forced:
        return forced
    if sys.platform == "darwin":
        try:
            import ocrmac  # noqa: F401
            return "vision"
        except ImportError:
            pass
    return "tesseract"


def ocr_page_lines(page) -> list[dict]:
    """OCR one page. Returns lines top-to-bottom: {text, x0, y0, x1, y1} in 0..1,
    origin top-left."""
    from PIL import Image

    # Greyscale keeps a 300-dpi page to ~8 MB instead of ~25 MB; small servers need that.
    pix = page.get_pixmap(dpi=OCR_DPI, alpha=False, colorspace=pymupdf.csGRAY)
    img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
    del pix
    try:
        lines = _ocr_vision(img) if _ocr_engine() == "vision" else _ocr_tesseract(img)
    finally:
        img.close()
    return _reading_order(lines)


def _ocr_vision(img) -> list[dict]:
    from ocrmac import ocrmac

    results = ocrmac.OCR(img, recognition_level="accurate",
                         language_preference=["en-US"]).recognize()
    # Vision's origin is bottom-left.
    return [{"text": text, "x0": x, "x1": x + w, "y0": 1 - (y + h), "y1": 1 - y}
            for text, _conf, (x, y, w, h) in results]


def _ocr_tesseract(img) -> list[dict]:
    import pytesseract

    d = pytesseract.image_to_data(img, lang=os.environ.get("OCR_LANG", "eng"),
                                  output_type=pytesseract.Output.DICT)
    W, H = img.size
    grouped: dict[tuple, dict] = {}
    for i, word in enumerate(d["text"]):
        if not word.strip() or float(d["conf"][i]) < 0:
            continue
        key = (d["block_num"][i], d["par_num"][i], d["line_num"][i])
        x0, y0 = d["left"][i], d["top"][i]
        x1, y1 = x0 + d["width"][i], y0 + d["height"][i]
        g = grouped.setdefault(key, {"words": [], "x0": x0, "y0": y0, "x1": x1, "y1": y1})
        g["words"].append(word)
        g.update(x0=min(g["x0"], x0), y0=min(g["y0"], y0),
                 x1=max(g["x1"], x1), y1=max(g["y1"], y1))
    return [{"text": " ".join(g["words"]), "x0": g["x0"] / W, "x1": g["x1"] / W,
             "y0": g["y0"] / H, "y1": g["y1"] / H} for g in grouped.values()]


def _reading_order(lines: list[dict]) -> list[dict]:
    """Sort into rows (top to bottom), left to right within a row."""
    lines = sorted(lines, key=lambda l: (l["y0"] + l["y1"]) / 2)
    rows: list[list[dict]] = []
    for ln in lines:
        mid = (ln["y0"] + ln["y1"]) / 2
        if rows:
            last = rows[-1][-1]
            if last["y0"] <= mid <= last["y1"]:
                rows[-1].append(ln)
                continue
        rows.append([ln])
    out = []
    for row in rows:
        out.extend(sorted(row, key=lambda l: l["x0"]))
    return out


def _text_layer_lines(page) -> list[dict]:
    w, h = page.rect.width, page.rect.height
    lines = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            text = "".join(s["text"] for s in line["spans"]).strip()
            if text:
                x0, y0, x1, y1 = line["bbox"]
                lines.append({"text": text, "x0": x0 / w, "x1": x1 / w,
                              "y0": y0 / h, "y1": y1 / h})
    return lines


def _paragraphs(lines: list[dict]) -> list[str]:
    """Merge consecutive lines into paragraphs using vertical gaps."""
    if not lines:
        return []
    heights = [l["y1"] - l["y0"] for l in lines]
    typical = statistics.median(heights)
    paras, current = [], [lines[0]]
    for prev, ln in zip(lines, lines[1:]):
        gap = ln["y0"] - prev["y1"]
        same_row = ln["y0"] < prev["y1"] and ln["x0"] > prev["x0"]
        new_para = (
            not same_row
            and (gap > typical * 0.7                         # visible blank space
                 or ln["y0"] < prev["y0"] - typical           # jumped back up (new column)
                 or prev["text"].rstrip().endswith((".", ":", "?", "!"))
                 and prev["x1"] < 0.75)                       # short last line
        )
        if new_para:
            paras.append(current)
            current = [ln]
        else:
            current.append(ln)
    paras.append(current)

    out = []
    for group in paras:
        text = ""
        for ln in group:
            t = ln["text"].strip()
            if text.endswith("-") and t[:1].islower():
                text = text[:-1] + t          # re-join a hyphenated line break
            else:
                text = f"{text} {t}".strip()
        out.append(text)
    return out


def _convert_ocr(pdf_path, docx_path, password, scanned: set[int], progress,
                 ocr_cache: dict | None = None) -> None:
    doc = pymupdf.open(pdf_path)
    if doc.needs_pass:
        doc.authenticate(password or "")
    out = Document()
    style = out.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    try:
        for page in doc:
            try:
                if page.number in scanned:
                    lines = ocr_page_lines(page)
                    if ocr_cache is not None:
                        ocr_cache[page.number] = lines
                else:
                    lines = _text_layer_lines(page)
            except Exception as exc:
                raise ConversionError(f"Text recognition failed on page "
                                      f"{page.number + 1} ({exc}).") from exc
            for para in _paragraphs(lines):
                out.add_paragraph(para)
            if page.number < doc.page_count - 1:
                out.add_page_break()
            if progress:
                progress(page.number + 1, doc.page_count, "Reading page")
    finally:
        doc.close()
    out.save(docx_path)
