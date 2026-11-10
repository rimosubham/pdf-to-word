"""Generate test PDFs: report with tables, two-column article, long document, scanned copy."""
import os
import random
import sys

import pymupdf
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (BaseDocTemplate, Frame, PageBreak, PageTemplate, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)
S = getSampleStyleSheet()
random.seed(7)

VOCAB = ("quarterly revenue growth margin customer retention pipeline forecast audit compliance "
         "infrastructure latency throughput deployment migration schedule vendor contract renewal "
         "analysis baseline variance threshold allocation budget headcount roadmap milestone review "
         "the of and to in for with on by from that this which was were has have will can should").split()


def sentence(n=None):
    n = n or random.randint(9, 22)
    words = [random.choice(VOCAB) for _ in range(n)]
    if random.random() < 0.5:
        words.insert(random.randrange(n), f"{random.randint(1, 9999):,}")
    if random.random() < 0.3:
        words.insert(random.randrange(n), f"{random.uniform(0.1, 99):.1f}%")
    return " ".join(words).capitalize() + "."


def para(k=None):
    return " ".join(sentence() for _ in range(k or random.randint(3, 6)))


def report():
    doc = SimpleDocTemplate(f"{OUT}/1_report_tables.pdf", pagesize=letter)
    el = [Paragraph("Annual Operations Report 2026", S["Title"]),
          Paragraph("Prepared for the Board — Confidential “draft” — naïve café résumé", S["Italic"]),
          Spacer(1, 12)]
    for sec in range(1, 6):
        el.append(Paragraph(f"{sec}. Section heading number {sec}", S["Heading1"]))
        el.append(Paragraph(para(), S["BodyText"]))
        el.append(Paragraph(f"{sec}.1 Sub-section with <b>bold</b>, <i>italic</i> and <u>underlined</u> text", S["Heading2"]))
        el.append(Paragraph(para(), S["BodyText"]))
        for i in range(3):
            el.append(Paragraph(f"• Bullet point {i + 1}: {sentence(8)}", S["BodyText"]))
        data = [["Region", "Q1", "Q2", "Q3", "Q4", "Total"]]
        for region in ["North America", "Europe", "Asia Pacific", "Latin America"]:
            q = [random.randint(1000, 99999) for _ in range(4)]
            data.append([region] + [f"${v:,}" for v in q] + [f"${sum(q):,}"])
        t = Table(data, hAlign="LEFT")
        t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                               ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                               ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")]))
        el += [Spacer(1, 8), t, Spacer(1, 12)]
    doc.build(el)


def two_column():
    path = f"{OUT}/2_two_column_article.pdf"
    doc = BaseDocTemplate(path, pagesize=A4)
    w, h = A4
    gap, m = 18, 54
    cw = (w - 2 * m - gap) / 2
    frames = [Frame(m, m, cw, h - 2 * m, id="l"), Frame(m + cw + gap, m, cw, h - 2 * m, id="r")]
    doc.addPageTemplates(PageTemplate(id="two", frames=frames))
    el = [Paragraph("A Study of Throughput Under Variable Load", S["Heading1"])]
    for i in range(1, 9):
        el.append(Paragraph(f"{i}. Topic {i}", S["Heading3"]))
        for _ in range(3):
            el.append(Paragraph(para(), S["BodyText"]))
    doc.build(el)


def long_doc():
    doc = SimpleDocTemplate(f"{OUT}/3_long_40_pages.pdf", pagesize=letter)
    el = []
    for ch in range(1, 21):
        el.append(Paragraph(f"Chapter {ch}", S["Heading1"]))
        for _ in range(9):
            el.append(Paragraph(para(6), S["BodyText"]))
        el.append(PageBreak())
    doc.build(el)


def scanned(src, dst, dpi=200):
    """Rasterise a PDF into an image-only PDF (no text layer), like a scanner."""
    s, d = pymupdf.open(src), pymupdf.open()
    for page in list(s)[:3]:
        pix = page.get_pixmap(dpi=dpi)
        p = d.new_page(width=page.rect.width, height=page.rect.height)
        p.insert_image(p.rect, pixmap=pix)
    d.save(dst)


report()
two_column()
long_doc()
scanned(f"{OUT}/1_report_tables.pdf", f"{OUT}/4_scanned_report.pdf")
print(sorted(os.listdir(OUT)))
