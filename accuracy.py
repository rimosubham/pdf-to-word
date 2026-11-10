"""Measure how faithfully the Word file reproduces the PDF's text.

The score is computed on every conversion, from the two actual files:
  * words in the PDF   (visible text layer; OCR for image-only pages)
  * words in the .docx (body, tables, headers, footers, text boxes)

accuracy      = character-level F1: drops when characters are lost, changed or
                invented; differences in spacing / line wrapping are ignored.
word_accuracy = stricter: share of whole words reproduced exactly. `order` additionally checks the words come out in the
same reading sequence.

Diagonal text (stamps such as a "DRAFT" watermark across the page) is not
part of the document's content; it is left out of the Word file and out of
the score, and reported separately.
"""
from __future__ import annotations

import re
import unicodedata
import warnings
import zipfile
from collections import Counter
from difflib import SequenceMatcher
from xml.etree import ElementTree as ET

warnings.filterwarnings("ignore")
import pymupdf  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
ORDER_CHECK_MAX_WORDS = 15000


def normalize(text: str) -> list[str]:
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("­", "")                    # soft hyphen
    text = re.sub(r"[‐-―−]", "-", text)     # dash variants
    text = re.sub(r"[‘’‛]", "'", text)
    text = re.sub(r"[“”]", '"', text)
    return re.findall(r"\d+(?:[.,]\d+)*|[^\W\d_]+", text.lower())


def open_pdf(pdf_path: str, password: str | None = None):
    doc = pymupdf.open(pdf_path)
    if doc.needs_pass:
        doc.authenticate(password or "")
    return doc


def visible_lines(page) -> list[dict]:
    """Visible text lines in the PDF's own block order.

    Each: {text, size, diagonal}. Text drawn invisibly (the hidden layer a
    scanner adds behind a page image) is skipped.
    """
    out = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            spans = [s for s in line["spans"] if s.get("alpha", 255) > 0]
            text = _join_spans(spans).strip()
            if not text:
                continue
            dx, dy = line.get("dir", (1.0, 0.0))
            out.append({
                "text": text,
                "size": max(s["size"] for s in spans),
                "diagonal": abs(dx) > 0.05 and abs(dy) > 0.05,
            })
    return out


def _join_spans(spans) -> str:
    """Concatenate a line's spans, adding a space where there is a visible gap
    (adjacent table cells often arrive as touching spans with no space)."""
    text, prev = "", None
    for sp in spans:
        if prev is not None and text and not text[-1].isspace() and not sp["text"][:1].isspace():
            gap = sp["bbox"][0] - prev["bbox"][2]
            if gap > 0.2 * max(sp["size"], prev["size"]):
                text += " "
        text += sp["text"]
        prev = sp
    return text


def pdf_lines(pdf_path: str, password: str | None = None,
              scanned_pages: set[int] | None = None) -> list[dict]:
    """Every line of the PDF, in order: {page, text, size, diagonal}."""
    doc = open_pdf(pdf_path, password)
    lines = []
    try:
        for page in doc:
            if scanned_pages and page.number in scanned_pages:
                from converter import ocr_page_lines
                found = [{"text": l["text"], "size": 11, "diagonal": False}
                         for l in ocr_page_lines(page)]
            else:
                found = visible_lines(page)
            for ln in found:
                ln["page"] = page.number
            lines.extend(found)
    finally:
        doc.close()
    return lines


def docx_paragraph_texts(root) -> list[str]:
    parts: list[str] = []
    _collect(root, parts)
    return parts


def docx_text(docx_path: str) -> str:
    parts: list[str] = []
    with zipfile.ZipFile(docx_path) as z:
        names = [n for n in z.namelist()
                 if re.fullmatch(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml", n)]
        for name in sorted(names, key=lambda n: (not n.endswith("document.xml"), n)):
            _collect(ET.fromstring(z.read(name)), parts)
    return "\n".join(parts)


def paragraph_text(p) -> str:
    buf = []
    for el in p.iter():
        if el.tag == f"{W}t":
            buf.append(el.text or "")
        elif el.tag in (f"{W}tab", f"{W}br"):
            buf.append(" ")
    return "".join(buf)


def _collect(node, parts: list[str]) -> None:
    """Paragraph text in document order, skipping legacy duplicates of text boxes."""
    for child in node:
        if child.tag == MC_FALLBACK:
            continue
        if child.tag == f"{W}p":
            parts.append(paragraph_text(child))
        else:
            _collect(child, parts)


def _char_accuracy(src: list[str], out: list[str], common: Counter) -> float:
    """Character-level F1. Words that match exactly count in full; the leftovers
    on each side are aligned character by character, so a word that was merely
    split or joined ("pre existing" / "preexisting") is not counted as lost."""
    def leftovers(words):
        budget, rest = Counter(common), []
        for w in words:
            if budget[w] > 0:
                budget[w] -= 1
            else:
                rest.append(w)
        return "".join(rest)

    a, b = leftovers(src), leftovers(out)
    exact = sum(len(w) * n for w, n in common.items())
    partial = 0
    if a and b and len(a) * len(b) <= 30_000_000:
        partial = sum(m.size for m in SequenceMatcher(None, a, b, autojunk=False).get_matching_blocks())
    matched = exact + partial
    n_src, n_out = exact + len(a), exact + len(b)
    if not matched:
        return 0.0
    r, p = matched / n_src, matched / n_out
    return 2 * r * p / (r + p)


def compare(pdf_path: str, docx_path: str, password: str | None = None,
            scanned_pages: set[int] | None = None, lines: list[dict] | None = None) -> dict:
    if lines is None:
        lines = pdf_lines(pdf_path, password, scanned_pages)
    content = [l["text"] for l in lines if not l["diagonal"]]
    watermark_words = sum(len(normalize(l["text"])) for l in lines if l["diagonal"])

    src = normalize("\n".join(content))
    out = normalize(docx_text(docx_path))
    if not src:
        return {"accuracy": None, "source_words": 0, "output_words": len(out),
                "missing": [], "extra": [], "order": None, "watermark_words": watermark_words,
                "note": "No text was found in the PDF, so accuracy cannot be measured."}

    cs, co = Counter(src), Counter(out)
    matched = sum((cs & co).values())
    recall = matched / len(src)
    precision = matched / len(out) if out else 0.0
    f1 = 2 * recall * precision / (recall + precision) if matched else 0.0

    char = _char_accuracy(src, out, cs & co)

    order = None
    if len(src) <= ORDER_CHECK_MAX_WORDS:
        order = SequenceMatcher(None, src, out, autojunk=False).ratio()

    return {
        "accuracy": round(char * 100, 2),
        "word_accuracy": round(f1 * 100, 2),
        "recall": round(recall * 100, 2),
        "precision": round(precision * 100, 2),
        "order": round(order * 100, 2) if order is not None else None,
        "source_words": len(src),
        "output_words": len(out),
        "watermark_words": watermark_words,
        "missing": [w for w, _ in (cs - co).most_common(12)],
        "extra": [w for w, _ in (co - cs).most_common(12)],
    }
