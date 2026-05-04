"""
user_data_compiler.py

将 data/mock_user_data.json 中的用户指令序列
通过 docx_tools.py 编译为 data/user_extraction.json。

输出格式与 data/extraction.json 完全一致:
    source          — 固定为 "template.docx"
    headers         — 从 extraction.json 原样复制
    footers         — 从 extraction.json 原样复制
    relationships   — 从 extraction.json 原样复制
    sections        — 从 body_elements 的 section_break 字段派生
    body_elements   — 由 docx_tools 函数序列生成

Usage:
    python user_data_compiler.py [user_data.json] [output.json]

Defaults:
    user_data.json  = data/mock_user_data.json
    output.json     = data/user_extraction.json
"""

import json
import sys
from pathlib import Path

try:
    from . import docx_tools as dt
    from . import base_agent as ba
except:
    import docx_tools as dt
    import base_agent as ba
# ── Paths ──────────────────────────────────────────────────────────────────────
_BASE          = Path(__file__).parent.parent
_EXTRACTION    = _BASE / "data" / "extraction.json"

# ── 函数名 → docx_tools 函数 映射 ────────────────────────────────────────────
_DISPATCH: dict = {
    "add_heading":                   dt.add_heading,
    "add_paragraph":                 dt.add_paragraph,
    "generate_toc":                  dt.generate_toc,
    "insert_page_break":             dt.insert_page_break,
    "insert_section_break":          dt.insert_section_break,
    "insert_figure":                 dt.insert_figure,
    "insert_table":                  dt.insert_table,
    "insert_equation":               dt.insert_equation,
    "insert_abstract_with_keywords": dt.insert_abstract_with_keywords,
    "add_reference":                 dt.add_reference,
}


def _load_extraction_scaffold() -> dict:
    """
    Load headers, footers, relationships from extraction.json.
    These fields are copied verbatim into the output so that
    docx_compiler.py can resolve rIds without any template scanning.
    Returns an empty scaffold if extraction.json is missing.
    """
    if not _EXTRACTION.exists():
        print(f"[WARN] {_EXTRACTION} not found — headers/footers/relationships will be empty.")
        return {"headers": {}, "footers": {}, "relationships": {}}

    with open(_EXTRACTION, encoding="utf-8") as f:
        ext: dict = json.load(f)

    return {
        "headers":       ext.get("headers",       {}),
        "footers":       ext.get("footers",        {}),
        "relationships": ext.get("relationships",  {}),
    }


def _derive_sections(body_elements: list[dict]) -> list[dict]:
    """
    Build the sections list (same format as extraction.json 'sections')
    by scanning body_elements for paragraphs whose section_break is not null.
    """
    sections: list[dict] = []
    for elem in body_elements:
        sb = elem.get("section_break")
        if sb is None:
            continue
        entry: dict = {
            "paragraph_index": elem.get("index"),
            "header_refs":     sb.get("header_refs", {}),
            "footer_refs":     sb.get("footer_refs",  {}),
            "page_size":       sb.get("page_size", {"w": "11906", "h": "16838"}),
        }
        if "restart_page_number" in sb:
            entry["restart_page_number"] = sb["restart_page_number"]
        sections.append(entry)
    return sections


def _collect_context(body_elements: list[dict], idx: int, window: int = 5) -> list[str]:
    texts: list[str] = []
    start = max(0, idx - window)
    end   = min(len(body_elements), idx + window + 1)
    for i in range(start, end):
        if i == idx:
            continue
        elem = body_elements[i]
        if elem.get('type') == 'paragraph':
            text = (elem.get('text') or '').strip()
            if text:
                texts.append(text)
    return texts


def _generate_caption(context_texts: list[str]) -> str:
    system_prompt = "你是一名学术论文排版助手，擅长为图片添加规范的中文图题。"
    if context_texts:
        ctx = '\n'.join(f'- {t}' for t in context_texts)
        user_prompt = (
            f"以下是文档中某张图片前后的文本段落：\n{ctx}\n\n"
            "请根据上下文为这张图片起一个简洁的中文图题（格式如'图X-X xxx示意图'）。"
            "如果上下文中已有明确的图题描述则直接提取，否则根据语义自行命名。"
            "只返回图题文本，不要任何解释。"
        )
    else:
        user_prompt = (
            "文档中有一张没有标题的图片，且没有可参考的上下文。"
            "请为其生成一个通用的中文图题（如'示意图'）。只返回图题文本，不要任何解释。"
        )
    try:
        return ba.call(system_prompt, user_prompt).strip()
    except Exception as exc:
        print(f"[WARN] caption LLM call failed: {exc}")
        return "示意图"


def fill_missing_captions(body_elements: list[dict]) -> None:
    """为 caption 为空的 image 元素调用 LLM 补全，并在其后插入 caption 段落。"""
    i = 0
    while i < len(body_elements):
        elem = body_elements[i]
        if elem.get('type') == 'image' and not elem.get('caption'):
            context  = _collect_context(body_elements, i)
            caption  = _generate_caption(context)
            elem['caption'] = caption
            cap_para = {
                "index":         -1,
                "type":          "paragraph",
                "style":         None,
                "text":          caption,
                "runs":          [dt._mk_run(caption, dt._CAPTION_RPR)],
                "pPr":           dt._CAPTION_PPR,
                "section_break": None,
            }
            body_elements.insert(i + 1, cap_para)
            i += 2
        else:
            i += 1


def compile_user_data(
    user_data_path: str = "data/mock_user_data.json",
    output_path:    str = "data/user_extraction.json",
) -> None:
    src = Path(user_data_path)
    if not src.exists():
        print(f"[ERROR] user data not found: {src}", file=sys.stderr)
        sys.exit(1)

    with open(src, encoding="utf-8") as f:
        user_data: dict = json.load(f)

    actions: list[dict] = user_data.get("document", [])
    if not actions:
        print("[WARN] 'document' list is empty — nothing to compile.")

    # ── Run docx_tools pipeline ────────────────────────────────────────────────
    dt.clear_document()

    skipped = 0
    for i, action in enumerate(actions):
        action = dict(action)               # shallow copy — don't mutate source
        fn_type = action.pop("type", None)

        if fn_type is None:
            print(f"[WARN] action[{i}] has no 'type' field — skipped.")
            skipped += 1
            continue

        fn = _DISPATCH.get(fn_type)
        if fn is None:
            print(f"[WARN] unknown action type '{fn_type}' at index {i} — skipped.")
            skipped += 1
            continue

        try:
            fn(**action)
        except TypeError as exc:
            print(f"[ERROR] action[{i}] type='{fn_type}': bad arguments — {exc}",
                  file=sys.stderr)
            skipped += 1

    body_elements = dt.get_document()
    fill_missing_captions(body_elements)

    # ── Assemble output JSON ───────────────────────────────────────────────────
    scaffold = _load_extraction_scaffold()

    output: dict = {
        "source":        "template.docx",
        "headers":       scaffold["headers"],
        "footers":       scaffold["footers"],
        "relationships": scaffold["relationships"],
        "sections":      _derive_sections(body_elements),
        "body_elements": body_elements,
    }

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"[OK] compiled {len(actions) - skipped} actions → {len(body_elements)} body_elements")
    print(f"     sections : {len(output['sections'])}")
    print(f"     headers  : {len(output['headers'])}")
    print(f"     footers  : {len(output['footers'])}")
    print(f"     output   : {out_path.resolve()}")
    if skipped:
        print(f"[WARN] {skipped} action(s) were skipped.")


if __name__ == "__main__":
    args = sys.argv[1:]
    compile_user_data(
        user_data_path = args[0] if len(args) > 0 else "docx_manager\\docx_engine\\data\\user_data.json",
        output_path    = args[1] if len(args) > 1 else "docx_manager\\docx_engine\\data\\user_extraction.json",
    )