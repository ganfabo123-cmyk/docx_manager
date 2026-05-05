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
import threading
import traceback
import uuid
from pathlib import Path
from typing import Literal
import sys

import tempfile
from pydantic import BaseModel

import pythoncom
import win32com.client as win32
import requests as http_requests
from flask import Flask, jsonify, request, send_file, abort, Response, stream_with_context

# ── Pipeline imports ───────────────────────────────────────────────────────────
# Set CWD to project root so that relative-path defaults inside engine modules
# (data/, templates/, sections_config/) resolve correctly.
# ── 路径修正：让 `docx_manager` 包和 `engine` 子包都可以被找到 ─────────────────
_BASE         = Path(__file__).parent          # docx_manager/docx_engine/
_PROJECT_ROOT = _BASE.parent.parent            # hit-paper-helper/

for _p in [str(_PROJECT_ROOT), str(_BASE)]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

os.chdir(_BASE)

from engine.docx_parser      import DocxParser            # step 1
from engine                  import user_data_generator as udg  # step 2
from engine                  import user_data_compiler  as udc  # step 3
from engine.docx_compiler    import DocxCompiler          # step 4
from engine                  import base_agent as ba

# ── WPS post-processing imports ────────────────────────────────────────────────
from docx_manager.wps_ui.workflows.hit_footer       import apply_hit_page_numbers
from docx_manager.wps_ui.workflows.insert_image     import insert_n_images_one_col
from docx_manager.wps_ui.workflows.insert_two_images import insert_n_images_two_col
from wps_com.insert_image import (
    insert_n_images_one_col as _com_one_col,
    insert_n_images_two_col as _com_two_col,
)



# ── Configuration ──────────────────────────────────────────────────────────────
_HIT_CONFIG    = str(_BASE / "sections_config" / "hit_config.json")
_TEMPLATE_DIR  = str(_BASE / "templates" / "hit-template")
_EXTRACTION    = str(_BASE / "data" / "extraction.json")
_OUTPUTS_DIR   = _BASE / "outputs"
_OUTPUTS_DIR.mkdir(exist_ok=True)
_ANCHOR_IMAGE  = str(_BASE.parent.parent / "anchor.png")   # ArUco 锚定图，服务端固定路径

_MAX_UPLOAD_MB = 32
_ALLOWED_EXT   = {".docx"}

# ── Per-job serialization locks ───────────────────────────────────────────────
_job_locks: dict[str, threading.Lock] = {}
_job_locks_mu = threading.Lock()

def _get_job_lock(job_id: str) -> threading.Lock:
    with _job_locks_mu:
        if job_id not in _job_locks:
            _job_locks[job_id] = threading.Lock()
        return _job_locks[job_id]


# ── Layout planning models ─────────────────────────────────────────────────────

class _ImageGroup(BaseModel):
    layout:   Literal["one_col", "two_col"]
    captions: list[str]

class _LayoutDecision(BaseModel):
    groups: list[_ImageGroup]


def _default_layout(captions: list[str]) -> list[dict]:
    """Fallback: every image gets its own single-column group."""
    return [{"layout": "one_col", "captions": [c]} for c in captions]


def _validate_layout(decision: _LayoutDecision, captions: list[str]) -> bool:
    """Check that all captions appear exactly once and group sizes are valid."""
    assigned: list[str] = []
    for g in decision.groups:
        if g.layout == "two_col" and len(g.captions) != 2:
            return False
        if g.layout == "one_col" and len(g.captions) != 1:
            return False
        assigned.extend(g.captions)
    return sorted(assigned) == sorted(captions)


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


def _get_output_docx(job_id: str) -> str | None:
    """Return the output.docx path for a job_id, or None if invalid / not found."""
    if not job_id or not all(c in "0123456789abcdef" for c in job_id):
        return None
    p = _OUTPUTS_DIR / job_id / "output.docx"
    return str(p) if p.exists() else None


def _lookup_deferred(job_id: str, captions: list[str]) -> tuple[list[dict], list[str]]:
    """
    Load deferred_images.json for job_id and return the subset matching captions
    (in the order requested).  Also returns a list of captions that were not found.

    Each returned item is the original deferred dict (contains file_path, anchor_text,
    width, height, etc.).
    """
    p = _OUTPUTS_DIR / job_id / "deferred_images.json"
    if not p.exists():
        return [], captions[:]
    with open(str(p), encoding="utf-8") as fh:
        all_deferred: list[dict] = json.load(fh)

    by_caption = {img["caption"]: img for img in all_deferred}
    found, missing = [], []
    for cap in captions:
        if cap in by_caption:
            found.append(by_caption[cap])
        else:
            missing.append(cap)
    return found, missing


def _refresh_ole_previews(docx_path: str) -> None:
    """
    Open the generated docx in WPS/Word so the OLE host renders all
    embedded equation objects and writes their WMF previews back into
    the document on save.

    Tries WPS (Kwps.Application) first, then Word (Word.Application).
    Silently skipped when neither application is available.
    """
    abs_path = os.path.abspath(docx_path)
    pythoncom.CoInitialize()
    app = doc = None
    used_prog_id = None
    try:
        for prog_id in ("Kwps.Application", "Word.Application"):
            try:
                app = win32.Dispatch(prog_id)
                used_prog_id = prog_id
                break
            except Exception:
                app = None
        if app is None:
            log.warning("[OLE] Neither WPS nor Word available — WMF previews skipped")
            return
        app.Visible = False
        app.DisplayAlerts = False
        doc = app.Documents.Open(abs_path)
        doc.Save()
        log.info("[OLE] Previews refreshed via %s", used_prog_id)
    except Exception as exc:
        log.warning("[OLE] Refresh failed (%s) — WMF previews may be missing", exc)
    finally:
        if doc:
            try: doc.Close(False)
            except Exception: pass
        if app:
            try: app.Quit()
            except Exception: pass
        pythoncom.CoUninitialize()


def _run_pipeline(input_docx: str, job_dir: Path) -> tuple[str, list[dict]]:
    """
    Execute all four pipeline stages for one job.

    Args:
        input_docx : absolute path to the uploaded .docx file
        job_dir    : dedicated scratch directory for this job

    Returns:
        (output_docx_path, deferred_images)
        deferred_images — list of {index, anchor_text, caption} dicts (server-side
        paths are stripped; clients use these to decide layout via /insert-image or /two-col)
    """
    full_parsed   = str(job_dir / "full_parsed.json")
    user_data     = str(job_dir / "user_data.json")
    user_extract  = str(job_dir / "user_extraction.json")
    output_docx   = str(job_dir / "output.docx")

    # ── Step 1: parse uploaded docx → full_parsed.json ────────────────────────
    log.info("[1/5] Parsing %s", input_docx)
    parser = DocxParser(input_docx)
    parser.parse()
    parser.to_json(full_parsed)
    log.info("      → %s", full_parsed)

    # ── Step 2: generate user_data.json ───────────────────────────────────────
    log.info("[2/5] Generating user_data …")
    udg.generate(
        full_parsed_path = full_parsed,
        config_path      = _HIT_CONFIG,
        output_path      = user_data,
    )
    log.info("      → %s", user_data)

    # ── Step 3: compile user_data → user_extraction.json ──────────────────────
    log.info("[3/5] Compiling user_data …")
    udc.compile_user_data(
        user_data_path = user_data,
        output_path    = user_extract,
    )
    log.info("      → %s", user_extract)

    # ── Step 4: compile docx (skip_images=True → images deferred for WPS UI) ──
    log.info("[4/5] Building output.docx …")
    compiler = DocxCompiler(
        extraction_path = user_extract,
        template_dir    = _TEMPLATE_DIR,
    )
    compiler.compile(output_path=output_docx, skip_images=True)
    log.info("      → %s  (%d deferred images)", output_docx, len(compiler.deferred_images))

    # ── Step 5: refresh OLE previews ─────────────────────────────────────────
    log.info("[5/5] Refreshing OLE previews …")
    # _refresh_ole_previews(output_docx)

    deferred = compiler.deferred_images
    with open(str(job_dir / "deferred_images.json"), "w", encoding="utf-8") as fh:
        json.dump(deferred, fh, ensure_ascii=False, indent=2)

    image_summary = [
        {
            "index":       i,
            "anchor_text": img["anchor_text"],
            "caption":     img["caption"],
        }
        for i, img in enumerate(deferred)
    ]
    return output_docx, image_summary


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
        output_path, image_summary = _run_pipeline(input_path, job_dir)
    except Exception as exc:
        log.error("Job %s failed:\n%s", job_id, traceback.format_exc())
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"status": "error", "message": str(exc)}), 500

    log.info("Job %s complete → %s  (%d images)", job_id, output_path, len(image_summary))
    return jsonify({
        "status":       "ok",
        "job_id":       job_id,
        "download_url": f"/download/{job_id}",
        "images":       image_summary,
    })


@app.route("/footer", methods=["POST"])
def footer():
    """
    Apply HIT page-number formatting to a compiled output.docx.

    Body: {"job_id": "<hex>", "body_section": 4}
    """
    data         = request.json or {}
    job_id       = data.get("job_id", "")
    body_section = int(data.get("body_section", 4))

    docx_path = _get_output_docx(job_id)
    if not docx_path:
        return jsonify({"status": "error", "message": "job not found"}), 404

    try:
        log.info("[footer] job=%s section=%d", job_id, body_section)
        apply_hit_page_numbers(docx_path, body_section)
        log.info("[footer] done")
    except Exception as exc:
        log.error("[footer] failed: %s", exc)
        return jsonify({"status": "error", "message": str(exc)}), 500

    return jsonify({"status": "ok", "download_url": f"/download/{job_id}"})


@app.route("/insert-image", methods=["POST"])
def insert_image():
    """
    Insert images (single-column layout) into a compiled output.docx.

    The server resolves image file paths, dimensions, and default anchor_text from the
    deferred_images saved during /convert.  The client only needs to name the captions.

    Body:
    {
        "job_id":      "<hex>",
        "captions":    ["图3-1 流程图", "图3-2 对比图"],
        "anchor_text": "3.1 实验"    // optional — overrides the stored anchor_text
    }
    """
    data               = request.json or {}
    job_id             = data.get("job_id", "")
    captions           = data.get("captions", [])
    original_captions  = data.get("original_captions") or captions   # lookup key
    anchor_text        = data.get("anchor_text") or None
    chapter            = int(data.get("chapter", 1))
    fig_start          = int(data.get("fig_start", 1))

    if not captions:
        return jsonify({"status": "error", "message": "captions required"}), 400
    if len(original_captions) != len(captions):
        return jsonify({"status": "error",
                        "message": "original_captions length must match captions"}), 400

    docx_path = _get_output_docx(job_id)
    if not docx_path:
        return jsonify({"status": "error", "message": "job not found"}), 404

    imgs, missing = _lookup_deferred(job_id, original_captions)
    if missing:
        return jsonify({"status": "error",
                        "message": f"captions not found in job: {missing}"}), 404

    resolved_anchor = anchor_text or imgs[0]["anchor_text"]
    images   = [img["file_path"] for img in imgs]
    width    = imgs[0].get("width")  or None
    height   = imgs[0].get("height") or None

    try:
        log.info("[insert-image] job=%s anchor=%r images=%d", job_id, resolved_anchor, len(images))
        insert_n_images_one_col(
            docx_path=docx_path,
            anchor_text=resolved_anchor,
            anchor_image=_ANCHOR_IMAGE,
            images=images,
            captions=captions,
            chapter=chapter,
            fig_start=fig_start,
        )
        log.info("[insert-image] done")
    except Exception as exc:
        log.error("[insert-image] failed: %s", exc)
        return jsonify({"status": "error", "message": str(exc)}), 500

    return jsonify({"status": "ok", "download_url": f"/download/{job_id}"})


@app.route("/two-col", methods=["POST"])
def two_col():
    """
    Insert images in a two-column layout in a compiled output.docx.

    The server resolves image paths and anchor_text from the deferred_images saved
    during /convert.  The client only needs captions (exactly 2) and optionally
    overrides for anchor_text and total_caption.

    Body:
    {
        "job_id":        "<hex>",
        "captions":      ["图3-3 子图a描述", "图3-4 子图b描述"],
        "anchor_text":   "3.3 误差分析",  // optional override
        "total_caption": "图3-3和图3-4",  // optional — auto-generated if omitted
        "debug":         false,
        "phases":        [1, 2, 3, 4, 5]
    }
    """
    data              = request.json or {}
    job_id            = data.get("job_id", "")
    captions          = data.get("captions", [])
    original_captions = data.get("original_captions") or captions   # lookup key
    anchor_text       = data.get("anchor_text") or None
    total_caption     = data.get("total_caption") or None
    debug             = bool(data.get("debug", False))
    phases            = tuple(int(p) for p in data.get("phases", [1, 2, 3, 4, 5]))

    if len(captions)%2 != 0:
        return jsonify({"status": "error", "message": "captions must contain exactly 2 items"}), 400
    if len(original_captions)%2 != 0:
        return jsonify({"status": "error",
                        "message": "original_captions must contain exactly 2 items"}), 400

    docx_path = _get_output_docx(job_id)
    if not docx_path:
        return jsonify({"status": "error", "message": "job not found"}), 404

    imgs, missing = _lookup_deferred(job_id, original_captions)
    if missing:
        return jsonify({"status": "error",
                        "message": f"captions not found in job: {missing}"}), 404

    resolved_anchor = anchor_text or imgs[0]["anchor_text"]
    images          = [img["file_path"] for img in imgs]
    resolved_total  = total_caption or f"{captions[0]}和{captions[1]}"

    try:
        log.info("[two-col] job=%s anchor=%r images=2", job_id, resolved_anchor)
        insert_n_images_two_col(
            docx_path=docx_path,
            anchor_text=resolved_anchor,
            anchor_image=_ANCHOR_IMAGE,
            images=images,
            captions=captions,
            total_caption=resolved_total,
            debug=debug,
            run_phases=phases,
        )
        log.info("[two-col] done")
    except Exception as exc:
        log.error("[two-col] failed: %s", exc)
        return jsonify({"status": "error", "message": str(exc)}), 500

    return jsonify({"status": "ok", "download_url": f"/download/{job_id}"})


@app.route("/plan-layout", methods=["POST"])
def plan_layout():
    """
    Execute ONE image-layout group per call to avoid client timeouts.

    First call: provide captions (+ optional user_instruction) — server plans
    all groups via LLM and executes the first one.

    Subsequent calls: provide groups (the remaining_groups from the previous
    response) — server executes groups[0] directly, no LLM involved.

    Body:
    {
        "job_id":           "<hex>",
        "user_instruction": "把图3-1和图3-2并排放，其他单独放",  // first call only
        "captions":         ["图3-1 流程图", "图3-2 对比图"],     // first call only
        "groups":           [...],   // subsequent calls: remaining_groups from last response
        "chapter":          3,
        "fig_start":        1        // pass fig_next from the previous response each time
    }

    Response (more groups remain):
      {"status": "progress", "executed": {"layout": "...", "captions": [...]},
       "remaining_groups": [...], "fig_next": N, "download_url": "/download/<job_id>"}

    Response (last group done):
      {"status": "ok", "download_url": "/download/<job_id>"}

    Response (error):
      {"status": "error", "message": "..."}
    """
    data        = request.json or {}
    job_id      = data.get("job_id", "")
    instruction = (data.get("user_instruction") or "").strip()
    captions    = data.get("captions", [])
    groups      = data.get("groups")          # None if client didn't pass it
    chapter     = int(data.get("chapter",   1))
    fig_start   = int(data.get("fig_start", 1))
    use_com     = bool(data.get("use_com", False))

    docx_path = _get_output_docx(job_id)
    if not docx_path:
        return jsonify({"status": "error", "message": "job not found"}), 404

    _groups_file = _OUTPUTS_DIR / job_id / "layout_groups.json"

    with _get_job_lock(job_id):

        # ── Resolve groups ─────────────────────────────────────────────────────────
        # Priority: client-supplied > server-saved > plan from scratch
        if groups is None:
            if _groups_file.exists():
                with open(str(_groups_file), encoding="utf-8") as fh:
                    groups = json.load(fh)
                log.info("[plan-layout] loaded %d groups from saved state", len(groups))
            elif not captions:
                return jsonify({"status": "error",
                                "message": "captions required on first call"}), 400
            elif not instruction:
                groups = _default_layout(captions)
            else:
                caption_list = "\n".join(f"{i + 1}. {c}" for i, c in enumerate(captions))
                system_prompt = (
                    "你是一名学术论文排版助手。根据用户说明将图片分组并指定排版方式。\n"
                    "规则：\n"
                    "- one_col（单列）：每组恰好 1 张图片，居中独占一行\n"
                    "- two_col（双列）：每组恰好 2 张图片，左右并排\n"
                    "- 每张图片必须出现在且仅出现在一个组中\n"
                    "- 用户未提及的图片默认归入 one_col 单独一组"
                )
                user_prompt = (
                    f"图片列表：\n{caption_list}\n\n"
                    f"用户排版说明：{instruction}\n\n"
                    "请按要求输出分组结果。"
                )
                try:
                    decision: _LayoutDecision = ba.call_structured(
                        system_prompt, user_prompt, _LayoutDecision
                    )
                    if not _validate_layout(decision, captions):
                        log.warning("[plan-layout] LLM output failed validation — using default")
                        groups = _default_layout(captions)
                    else:
                        groups = [g.model_dump() for g in decision.groups]
                except Exception as exc:
                    log.warning("[plan-layout] LLM call failed (%s) — using default", exc)
                    groups = _default_layout(captions)

        if not groups:
            if _groups_file.exists():
                return jsonify({"status": "ok", "download_url": f"/download/{job_id}"})
            return jsonify({"status": "error", "message": "no groups to process"}), 400

        log.info("[plan-layout] job=%s groups_remaining=%d fig_start=%d",
                 job_id, len(groups), fig_start)

        # ── COM 模式：一次性处理所有组 ────────────────────────────────────────────────
        if use_com:
            cur_fig = fig_start
            try:
                for g in groups:
                    g_captions = g["captions"]
                    g_layout   = g["layout"]

                    g_imgs, g_missing = _lookup_deferred(job_id, g_captions)
                    if g_missing:
                        return jsonify({"status": "error",
                                        "message": f"captions not found: {g_missing}"}), 404

                    g_anchor = g_imgs[0]["anchor_text"]
                    g_images = [img["file_path"] for img in g_imgs]

                    if g_layout == "one_col":
                        log.info("[plan-layout/com] one_col anchor=%r fig=%d", g_anchor, cur_fig)
                        _com_one_col(
                            docx_path=docx_path,
                            anchor_text=g_anchor,
                            images=g_images,
                            captions=g_captions,
                            chapter=chapter,
                            fig_start=cur_fig,
                        )
                    else:  # two_col
                        total_caption = f"{g_captions[0]}和{g_captions[1]}"
                        log.info("[plan-layout/com] two_col anchor=%r", g_anchor)
                        _com_two_col(
                            docx_path=docx_path,
                            anchor_text=g_anchor,
                            images=g_images,
                            captions=g_captions,
                            total_caption=total_caption,
                        )

                    cur_fig += len(g_captions)

            except Exception as exc:
                log.error("[plan-layout/com] failed: %s", exc)
                return jsonify({"status": "error", "message": str(exc)}), 500

            if _groups_file.exists():
                _groups_file.unlink()

            return jsonify({"status": "ok", "download_url": f"/download/{job_id}"})

        # ── UI 模式：每次只处理第一组，客户端轮询 ────────────────────────────────────
        group          = groups[0]
        remaining      = groups[1:]
        group_captions = group["captions"]
        layout         = group["layout"]

        imgs, missing = _lookup_deferred(job_id, group_captions)
        if missing:
            return jsonify({"status": "error",
                            "message": f"captions not found: {missing}"}), 404

        anchor = imgs[0]["anchor_text"]
        images = [img["file_path"] for img in imgs]

        try:
            if layout == "one_col":
                log.info("[plan-layout] one_col anchor=%r fig=%d", anchor, fig_start)
                insert_n_images_one_col(
                    docx_path=docx_path,
                    anchor_text=anchor,
                    anchor_image=_ANCHOR_IMAGE,
                    images=images,
                    captions=group_captions,
                    chapter=chapter,
                    fig_start=fig_start,
                )
            else:  # two_col
                total_caption = f"{group_captions[0]}和{group_captions[1]}"
                log.info("[plan-layout] two_col anchor=%r", anchor)
                insert_n_images_two_col(
                    docx_path=docx_path,
                    anchor_text=anchor,
                    anchor_image=_ANCHOR_IMAGE,
                    images=images,
                    captions=group_captions,
                    total_caption=total_caption,
                )
        except Exception as exc:
            log.error("[plan-layout] failed: %s", exc)
            return jsonify({"status": "error", "message": str(exc)}), 500

        # ── Persist remaining groups so client doesn't have to carry state ────────
        with open(str(_groups_file), "w", encoding="utf-8") as fh:
            json.dump(remaining, fh, ensure_ascii=False)

        return jsonify({
            "status":           "ok" if not remaining else "progress",
            "executed":         {"layout": layout, "captions": group_captions},
            "remaining_groups": remaining,
            "fig_next":         fig_start + len(group_captions),
            "download_url":     f"/download/{job_id}",
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

