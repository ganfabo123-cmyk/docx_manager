"""
server.py

Flask HTTP server that accepts a .docx upload, runs the full conversion
pipeline, and returns a download link for the reformatted document.

Pipeline
--------
  1. DocxParser          : uploaded .docx  →  full_parsed.json
  2. user_data_generator : full_parsed.json + hit_config.json  →  user_data.json
  3. user_data_compiler  : user_data.json  →  user_extraction.json
  4. DocxCompiler        : user_extraction.json + template  →  output.docx

Endpoints
---------
  POST /convert
      Form field : file  (multipart/form-data, .docx)
      Response   : {"job_id": "...", "download_url": "/download/<job_id>"}

  GET  /download/<job_id>
      Response   : output.docx as attachment

  GET  /health
      Response   : {"status": "ok"}

Usage
-----
  python server.py [--host 0.0.0.0] [--port 5000]
"""

import json
import logging
import os
import shutil
import traceback
import uuid
from pathlib import Path

import tempfile

import pythoncom
import win32com.client as win32
import requests as http_requests
from flask import Flask, jsonify, request, send_file, abort

# ── Pipeline imports ───────────────────────────────────────────────────────────
# Set CWD to project root so that relative-path defaults inside engine modules
# (data/, templates/, sections_config/) resolve correctly.
_BASE = Path(__file__).parent
os.chdir(_BASE)

from engine.docx_parser      import DocxParser            # step 1
from engine                  import user_data_generator as udg  # step 2
from engine                  import user_data_compiler  as udc  # step 3
from engine.docx_compiler    import DocxCompiler          # step 4

# ── Configuration ──────────────────────────────────────────────────────────────
_HIT_CONFIG   = str(_BASE / "sections_config" / "hit_config.json")
_TEMPLATE_DIR = str(_BASE / "templates" / "hit-template")
_EXTRACTION   = str(_BASE / "data" / "extraction.json")
_OUTPUTS_DIR  = _BASE / "outputs"
_OUTPUTS_DIR.mkdir(exist_ok=True)

_MAX_UPLOAD_MB = 32
_ALLOWED_EXT   = {".docx"}

# ── Flask app ──────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = _MAX_UPLOAD_MB * 1024 * 1024

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in _ALLOWED_EXT


def _run_pipeline(input_docx: str, job_dir: Path) -> str:
    """
    Execute all four pipeline stages for one job.

    Args:
        input_docx : absolute path to the uploaded .docx file
        job_dir    : dedicated scratch directory for this job

    Returns:
        Absolute path to the generated output .docx
    """
    full_parsed   = str(job_dir / "full_parsed.json")
    user_data     = str(job_dir / "user_data.json")
    user_extract  = str(job_dir / "user_extraction.json")
    output_docx   = str(job_dir / "output.docx")

    # ── Step 1: parse uploaded docx → full_parsed.json ────────────────────────
    log.info("[1/4] Parsing %s", input_docx)
    parser = DocxParser(input_docx)
    parser.parse()
    parser.to_json(full_parsed)
    log.info("      → %s", full_parsed)

    # ── Step 2: generate user_data.json ───────────────────────────────────────
    log.info("[2/4] Generating user_data …")
    udg.generate(
        full_parsed_path = full_parsed,
        config_path      = _HIT_CONFIG,
        output_path      = user_data,
    )
    log.info("      → %s", user_data)

    # ── Step 3: compile user_data → user_extraction.json ──────────────────────
    log.info("[3/4] Compiling user_data …")
    udc.compile_user_data(
        user_data_path = user_data,
        output_path    = user_extract,
    )
    log.info("      → %s", user_extract)

    # ── Step 4: compile docx ──────────────────────────────────────────────────
    log.info("[4/4] Building output.docx …")
    DocxCompiler(
        extraction_path = user_extract,
        template_dir    = _TEMPLATE_DIR,
    ).compile(output_path=output_docx)
    log.info("      → %s", output_docx)

    return output_docx


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/convert", methods=["POST"])
def convert():
    """
    Accept a JSON body with a file URL, run the pipeline, return a download link.

    Expects: application/json  {"url": "<docx_url>"}
    Also accepts: {"docx": {"url": "<docx_url>"}}

    Response: {"status": "ok", "download_url": "/download/<job_id>"}
    """
    data = request.json or {}
    log.info("Payload: %s", json.dumps(data, ensure_ascii=False)[:300])

    # Support {"url": "..."} and {"docx": {"url": "..."}}
    file_url = data.get("url", "")
    if not file_url and isinstance(data.get("docx"), dict):
        file_url = data["docx"].get("url", "")

    if not file_url:
        return jsonify({"status": "error", "message": "No URL provided"}), 400

    log.info("Downloading file from: %s", file_url)
    try:
        resp = http_requests.get(file_url, timeout=30)
        resp.raise_for_status()
        content = resp.content
    except Exception as exc:
        log.error("Download failed: %s", exc)
        return jsonify({"status": "error", "message": f"Failed to download file: {exc}"}), 400

    # ── 先把下载内容存为 .doc，再用 Word 转成 .docx ──────────────────────────────
    temp_dir  = tempfile.gettempdir()
    raw_path  = os.path.join(temp_dir, f"input_{os.getpid()}.doc")
    with open(raw_path, "wb") as fh:
        fh.write(content)
    log.info("Downloaded %d bytes → %s", len(content), raw_path)

    abs_raw  = os.path.abspath(raw_path)
    base     = os.path.splitext(os.path.basename(abs_raw))[0]
    abs_docx = os.path.abspath(os.path.join(temp_dir, f"{base}.docx"))

    log.info("Converting .doc → .docx via Word COM …")
    pythoncom.CoInitialize()
    word = doc = None
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(abs_raw)
        doc.SaveAs2(abs_docx, FileFormat=16)
        log.info("Word conversion OK → %s", abs_docx)
    except Exception as exc:
        log.error("Word conversion failed: %s", exc)
        return jsonify({"status": "error", "message": f"Word conversion failed: {exc}"}), 500
    finally:
        if doc:
            doc.Close(False)
        if word:
            word.Quit()
        pythoncom.CoUninitialize()

    if not os.path.exists(abs_docx):
        return jsonify({"status": "error", "message": "Word finished but no .docx produced"}), 500

    job_id  = uuid.uuid4().hex
    job_dir = _OUTPUTS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    input_path = str(job_dir / "input.docx")
    shutil.copy2(abs_docx, input_path)
    log.info("Job %s: input ready → %s", job_id, input_path)

    # 清理临时文件
    for p in (raw_path, abs_docx):
        try:
            os.remove(p)
        except OSError:
            pass

    try:
        output_path = _run_pipeline(input_path, job_dir)
    except Exception as exc:
        log.error("Job %s failed:\n%s", job_id, traceback.format_exc())
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"status": "error", "message": str(exc)}), 500

    log.info("Job %s complete → %s", job_id, output_path)
    return jsonify({
        "status":       "ok",
        "download_url": f"/download/{job_id}",
    })


@app.route("/download/<job_id>", methods=["GET"])
def download(job_id: str):
    """Serve the generated .docx for a completed job."""
    # Sanitise job_id: only hex characters allowed
    if not all(c in "0123456789abcdef" for c in job_id):
        abort(400)

    output_path = _OUTPUTS_DIR / job_id / "output.docx"
    if not output_path.exists():
        abort(404)

    return send_file(
        str(output_path),
        as_attachment=True,
        download_name="output.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ── Direct invocation (use main.py instead) ────────────────────────────────────

if __name__ == "__main__":
    import main
    main.main()
