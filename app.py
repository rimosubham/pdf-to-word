"""PDF to Word converter - web server.

Run:  .venv/bin/python app.py     then open http://127.0.0.1:5055
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import threading
import time
import uuid

from flask import Flask, abort, jsonify, request, send_file, send_from_directory

import converter

BASE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(tempfile.gettempdir(), "pdf-to-word-jobs")
MAX_MB = int(os.environ.get("MAX_MB", 100))
KEEP_SECONDS = 60 * 60          # converted files are deleted after an hour

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024

jobs: dict[str, dict] = {}
lock = threading.Lock()
convert_slot = threading.Semaphore(int(os.environ.get("MAX_PARALLEL", 2)))
os.makedirs(WORK, exist_ok=True)


def _cleanup() -> None:
    cutoff = time.time() - KEEP_SECONDS
    with lock:
        old = [j for j, d in jobs.items() if d["created"] < cutoff]
        for j in old:
            jobs.pop(j, None)
    for j in old:
        shutil.rmtree(os.path.join(WORK, j), ignore_errors=True)


def _run(job_id: str, pdf_path: str, docx_path: str, password: str | None) -> None:
    job = jobs[job_id]

    def progress(done, total, label="Working"):
        job["progress"] = f"{label} {done} of {total}" if done else label

    with convert_slot:
        started = time.time()
        try:
            job["status"] = "converting"
            result = converter.convert(pdf_path, docx_path, password, progress)
            # A scanned page is checked against a second OCR pass of the same
            # image, which cannot catch recognition mistakes.
            result["verified"] = result["scanned_pages"] == 0
            job.update(status="done", result={
                **result,
                "seconds": round(time.time() - started, 1),
                "size": os.path.getsize(docx_path),
            })
        except converter.ConversionError as exc:
            job.update(status="error", error=str(exc))
        except Exception as exc:  # noqa: BLE001
            app.logger.exception("conversion failed")
            job.update(status="error", error=f"Unexpected error: {exc}")
        finally:
            try:
                os.remove(pdf_path)       # never keep the upload
            except OSError:
                pass


@app.get("/api/config")
def config():
    return jsonify(max_mb=MAX_MB)


@app.get("/healthz")
def health():
    return "ok"


@app.get("/")
def index():
    return send_from_directory(os.path.join(BASE, "static"), "index.html")


@app.post("/api/convert")
def start():
    _cleanup()
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="Choose a PDF file first."), 400
    if f.stream.read(5) != b"%PDF-":
        return jsonify(error="That file is not a PDF."), 400
    f.stream.seek(0)

    job_id = uuid.uuid4().hex
    folder = os.path.join(WORK, job_id)
    os.makedirs(folder)
    pdf_path = os.path.join(folder, "input.pdf")
    docx_path = os.path.join(folder, "output.docx")
    f.save(pdf_path)

    stem = os.path.splitext(os.path.basename(f.filename))[0]
    stem = re.sub(r"[^\w\- .()]+", "_", stem).strip() or "converted"
    with lock:
        jobs[job_id] = {"status": "queued", "created": time.time(), "progress": "",
                        "name": f"{stem}.docx", "docx": docx_path}
    threading.Thread(target=_run, daemon=True,
                     args=(job_id, pdf_path, docx_path,
                           request.form.get("password") or None)).start()
    return jsonify(id=job_id)


@app.get("/api/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify(error="This conversion has expired. Upload the file again."), 404
    return jsonify({k: job.get(k) for k in ("status", "progress", "error", "result", "name")})


@app.get("/api/download/<job_id>")
def download(job_id):
    job = jobs.get(job_id)
    if not job or job["status"] != "done" or not os.path.exists(job["docx"]):
        abort(404)
    return send_file(
        job["docx"], as_attachment=True, download_name=job["name"],
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.errorhandler(413)
def too_big(_):
    return jsonify(error=f"That file is larger than {MAX_MB} MB."), 413


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 5055)), threaded=True)
