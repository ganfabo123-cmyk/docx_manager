"""
build.py — 本地一键生成 HIT 学位论文 .docx

Pipeline
--------
  1. DocxParser          : input.docx  →  full_parsed.json  (临时)
  2. user_data_generator : full_parsed.json  →  user_data.json  (临时)
  3. user_data_compiler  : user_data.json  →  user_extraction.json  (临时)
  4. DocxCompiler        : user_extraction.json  →  output.docx
                           skip_images=True → 图片不写入 XML，存为文件
  5. WPS UI 单列排版     : 前 N-2 张图逐一插入
  6. WPS UI 双列排版     : 末 2 张图并排（需 --anchor-image）
  7. apply_hit_page_numbers : 设置 HIT 页码格式

用法
----
  python build.py input.docx output.docx [--anchor-image PATH] [--body-section 4]

若 deferred_images 不足 3 张，或未提供 --anchor-image，全部走单列排版。
"""

import argparse
import os
import sys
import tempfile
from pathlib import Path

# ── 路径修正：让 `docx_manager` 包和 `engine` 子包都可以被找到 ─────────────────
_BASE         = Path(__file__).parent          # docx_manager/docx_engine/
_PROJECT_ROOT = _BASE.parent.parent            # hit-paper-helper/

for _p in [str(_PROJECT_ROOT), str(_BASE)]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

os.chdir(_BASE)

from engine.docx_parser   import DocxParser
from engine               import user_data_generator as udg
from engine               import user_data_compiler  as udc
from engine.docx_compiler import DocxCompiler

from docx_manager.wps_ui.workflows.hit_footer        import apply_hit_page_numbers
from docx_manager.wps_ui.workflows.insert_image      import insert_n_images_one_col
from docx_manager.wps_ui.workflows.insert_two_images import insert_n_images_two_col

_HIT_CONFIG   = str(_BASE / "sections_config" / "hit_config.json")
_TEMPLATE_DIR = str(_BASE / "templates"       / "hit-template")


# ── Pipeline ───────────────────────────────────────────────────────────────────

def run_pipeline(input_docx: str, output_docx: str) -> list[dict]:
    """
    Steps 1–4: parse → generate → compile (skip_images=True).

    Returns compiler.deferred_images — each item:
        {anchor_text, caption, file_path, drawing_xml, width, height, position}
    """
    with tempfile.TemporaryDirectory(prefix="hit_build_") as tmp:
        full_parsed  = os.path.join(tmp, "full_parsed.json")
        user_data    = os.path.join(tmp, "user_data.json")
        user_extract = os.path.join(tmp, "user_extraction.json")

        print("[1/4] Parsing …")
        parser = DocxParser(input_docx)
        parser.parse()
        parser.to_json(full_parsed)

        print("[2/4] Generating user_data …")
        udg.generate(
            full_parsed_path=full_parsed,
            config_path=_HIT_CONFIG,
            output_path=user_data,
        )

        print("[3/4] Compiling user_data …")
        udc.compile_user_data(
            user_data_path=user_data,
            output_path=user_extract,
        )

        print("[4/4] Building output.docx (skip_images=True) …")
        compiler = DocxCompiler(
            extraction_path=user_extract,
            template_dir=_TEMPLATE_DIR,
        )
        compiler.compile(output_path=output_docx, skip_images=True)

    return compiler.deferred_images


# ── Helpers ────────────────────────────────────────────────────────────────────

def _caption(img: dict, fallback_idx: int) -> str:
    """Return caption, generating a default like '图1' when empty."""
    return img["caption"] or f"图{fallback_idx + 1}"


# ── Image insertion ────────────────────────────────────────────────────────────

def insert_images(
    output_docx:  str,
    deferred:     list[dict],
    anchor_image: str,
    body_section: int,
    layout:       str,          # "single" | "two-col" | "both"
) -> None:
    """
    Steps 5–7: insert images via WPS UI, then apply HIT page numbers.

    layout="both"    → deferred[:-2] 单列, deferred[-2:] 双列 (需 len>=3)
    layout="single"  → 所有图单列，一次调用
    layout="two-col" → 末 2 张双列，其余忽略
    """
    if len(deferred) >= 3 and layout in ("both", "two-col"):
        single_items  = deferred[:-2] if layout == "both" else []
        two_col_items = deferred[-2:]
    else:
        single_items  = deferred
        two_col_items = []

    # ── 单列：一次调用传入所有图 ───────────────────────────────────────────────
    valid = [(i, img) for i, img in enumerate(single_items) if img.get("file_path")]
    if valid:
        first_anchor = valid[0][1]["anchor_text"]
        images   = [img["file_path"]       for _, img in valid]
        captions = [_caption(img, global_i) for global_i, img in valid]
        width    = valid[0][1]["width"]  or None
        height   = valid[0][1]["height"] or None

        print(f"[single-col] anchor={first_anchor!r}  images={len(images)}")
        insert_n_images_one_col(
            docx_path=output_docx,
            anchor_text=first_anchor,
            anchor_image=anchor_image,
            images=images,
            captions=captions,
        )

    skipped = len(single_items) - len(valid)
    if skipped:
        print(f"[skip] {skipped} 张 drawing_xml 图无 file_path，已略过")

    # ── 双列 ──────────────────────────────────────────────────────────────────
    if two_col_items:
        a_idx = len(deferred) - 2
        b_idx = len(deferred) - 1
        a, b  = two_col_items[0], two_col_items[1]
        a_path, b_path = a.get("file_path"), b.get("file_path")

        if not a_path or not b_path:
            print("[skip] 双列图含 drawing_xml，退为单列")
            for gi, img in [(a_idx, a), (b_idx, b)]:
                fp = img.get("file_path")
                if fp:
                    insert_n_images_one_col(
                        docx_path=output_docx,
                        anchor_text=img["anchor_text"],
                        anchor_image=anchor_image,
                        images=[fp],
                        captions=[_caption(img, gi)],
                    )
        else:
            cap_a = _caption(a, a_idx)
            cap_b = _caption(b, b_idx)
            total_caption = f"{cap_a}和{cap_b}"
            print(f"[two-col] anchor={a['anchor_text']!r}  images=2")
            insert_n_images_two_col(
                docx_path=output_docx,
                anchor_text=a["anchor_text"],
                anchor_image=anchor_image,
                images=[a_path, b_path],
                captions=[cap_a, cap_b],
                total_caption=total_caption,
            )

    # ── 页码 ──────────────────────────────────────────────────────────────────
    print(f"[footer] body_section={body_section}")
    apply_hit_page_numbers(output_docx, body_section)


# ── Entry point ────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description="HIT 学位论文一键生成工具")
    parser.add_argument("input_docx",  help="源 .docx 文件路径")
    parser.add_argument("output_docx", help="输出 .docx 文件路径")
    parser.add_argument(
        "--anchor-image",
        dest="anchor_image",
        default=r'D:\PycharmProjects\hit-paper-helper\anchor.png',
        metavar="PATH",
        help="双列排版用 ArUco 锚定图路径（不传则全部单列）",
    )
    parser.add_argument(
        "--body-section",
        dest="body_section",
        type=int,
        default=4,
        help="正文（绪论）所在节编号，默认 4",
    )
    parser.add_argument(
        "--layout",
        dest="layout",
        choices=["single", "two-col", "both"],
        default="both",
        help="图片排版模式：single 单列 / two-col 双列 / both 自动分配（默认）",
    )
    parser.add_argument(
        "--no-wps",
        dest="no_wps",
        action="store_true",
        default=False,
        help="只跑 compile，跳过 WPS UI 步骤（用于调试）",
    )
    args = parser.parse_args()

    input_docx  = os.path.abspath(args.input_docx)
    output_docx = os.path.abspath(args.output_docx)

    if not os.path.exists(input_docx):
        print(f"[ERROR] 找不到输入文件: {input_docx}", file=sys.stderr)
        sys.exit(1)

    Path(output_docx).parent.mkdir(parents=True, exist_ok=True)

    deferred = run_pipeline(input_docx, output_docx)
    print(f"\n生成完毕: {output_docx}")
    print(f"待插入图片: {len(deferred)} 张")
    for i, img in enumerate(deferred):
        print(f"  [{i}] anchor={img['anchor_text']!r}  caption={img['caption']!r}"
              f"  file={img['file_path']}")

    if args.no_wps:
        print("\n[--no-wps] 跳过 WPS UI 步骤，结束。")
        return

    if not deferred:
        print("\n无待插入图片，直接应用页码。")
        apply_hit_page_numbers(output_docx, args.body_section)
        return

    insert_images(output_docx, deferred, args.anchor_image, args.body_section, args.layout)
    print("\n全部完成。")


if __name__ == "__main__":
    main()
