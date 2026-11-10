"""Convert every PDF in a folder and print the measured accuracy.

usage: .venv/bin/python tests/run_accuracy.py <folder-with-pdfs> [--truth-for-scans]
"""
import glob
import os
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import converter  # noqa: E402

folder = sys.argv[1]
out_dir = tempfile.mkdtemp(prefix="p2w-")
rows = []
for pdf in sorted(glob.glob(os.path.join(folder, "*.pdf"))):
    name = os.path.basename(pdf)
    docx = os.path.join(out_dir, name[:-4] + ".docx")
    t = time.time()
    try:
        rep = converter.convert(pdf, docx)
        rows.append((name[:38], rep["engine"], rep["pages"], rep["source_words"],
                     rep["accuracy"], rep.get("word_accuracy"), rep.get("order"),
                     f"{time.time() - t:.1f}s  fixed={rep['repaired_lines']} wm={rep['watermark_words']}"))
    except Exception as exc:  # noqa: BLE001
        rows.append((name[:38], "FAILED", "-", "-", "-", "-", "-", str(exc)[:60]))

print(f"{'file':40}{'engine':8}{'pages':>6}{'words':>8}{'accuracy':>10}{'words%':>8}{'order':>8}  time")
for r in rows:
    print(f"{r[0]:40}{r[1]:8}{str(r[2]):>6}{str(r[3]):>8}{str(r[4]):>10}{str(r[5]):>8}{str(r[6]):>8}  {r[7]}")
print("output:", out_dir)
