# PDF to Word Converter

Local website that converts a PDF to an editable `.docx` and scores every
conversion against the original text.

## Run

    .venv/bin/python app.py

Open http://127.0.0.1:5055

## How it works

- `converter.py` — text PDFs go through pdf2docx (keeps tables, images, fonts,
  columns). Scanned PDFs (including scans with a hidden text layer) are read
  with Apple Vision OCR and rebuilt as paragraphs.
- `repair.py` — after conversion, any PDF line that did not reach the Word file
  is re-inserted next to its nearest neighbour.
- `accuracy.py` — character-level and word-level comparison of PDF vs docx.
- `app.py` / `static/index.html` — upload, progress, score, download.

## Test

    .venv/bin/python tests/make_samples.py /tmp/samples
    .venv/bin/python tests/run_accuracy.py /tmp/samples     # or any folder of PDFs

## Limits

- The score measures text, not visual layout.
- For scanned pages the score cannot detect a misread character.
- OCR uses macOS Vision, so the scanned-PDF path only runs on a Mac.
