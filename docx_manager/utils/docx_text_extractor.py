#!/usr/bin/env python3
"""
docx_text_extractor.py — 文本提取器
从 docx_parser.py 生成的 JSON 中提取所有文本内容，
输出为 [{id, content}] 格式的 JSON，供 LLM 进行样式分类。

用法:
    python docx_text_extractor.py input.json output_texts.json
    python docx_text_extractor.py input.json          # 自动命名
"""

import argparse
import json
import os
import sys

# ──────────────────────────── 文本提取工具 ────────────────────────────

def extract_runs_text(children):
    """从段落的 children 列表中拼出纯文本。"""
    parts = []
    for item in children:
        t = item.get("type", "")
        if t == "run":
            parts.append(item.get("text", ""))
        elif t == "hyperlink":
            for r in item.get("runs", []):
                parts.append(r.get("text", ""))
        elif t in ("ins",):          # 接受的修订，视作正文
            for r in item.get("runs", []):
                parts.append(r.get("text", ""))
        elif t == "fldChar":         # 域代码中的显示文本
            for r in item.get("runs", []):
                parts.append(r.get("text", ""))
        # del / drawing / bookmark 等不贡献纯文本
    return "".join(parts)


def extract_paragraph_text(para):
    """返回段落的纯文本（去掉两端空白）。"""
    return extract_runs_text(para.get("children", []))


def walk_blocks(blocks, results, path_prefix="body"):
    """
    递归遍历 body / 表格单元格 / 页眉页脚等 block 列表，
    为每个段落生成一条 {id, content} 记录。

    id 格式示例：
      body[0]                   → 顶层第 0 个 block（段落）
      body[3].row[1].cell[2].p[0]  → 表格内段落
    """
    for i, block in enumerate(blocks):
        btype = block.get("type", "")
        node_id = f"{path_prefix}[{i}]"

        if btype == "paragraph":
            text = extract_paragraph_text(block)
            # 保留空段落（可能是故意的间距段），但可选过滤
            results.append({"id": node_id, "content": text})

        elif btype == "table":
            for ri, row in enumerate(block.get("rows", [])):
                for ci, cell in enumerate(row.get("cells", [])):
                    cell_prefix = f"{node_id}.row[{ri}].cell[{ci}].p"
                    walk_blocks(cell.get("content", []), results, path_prefix=cell_prefix)

        elif btype == "sdt":
            # 结构化文档标签：如果已展开则递归
            inner = block.get("content", [])
            if inner:
                walk_blocks(inner, results, path_prefix=f"{node_id}.sdt")

        # drawing / unknown 等无文本，跳过


def extract_texts(data):
    """
    从完整 parsed JSON 中提取文本列表。
    同时处理：body、headers、footers。
    """
    results = []

    # 主体
    walk_blocks(data.get("body", []), results, path_prefix="body")

    # 页眉
    for rId, blocks in data.get("headers", {}).items():
        walk_blocks(blocks, results, path_prefix=f"header[{rId}]")

    # 页脚
    for rId, blocks in data.get("footers", {}).items():
        walk_blocks(blocks, results, path_prefix=f"footer[{rId}]")

    return results


# ──────────────────────────── CLI ─────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="从 docx_parser JSON 提取纯文本列表")
    parser.add_argument("input",  help="输入 .json 文件（docx_parser 输出）")
    parser.add_argument("output", nargs="?", help="输出文本 .json 文件（默认 *_texts.json）")
    parser.add_argument("--keep-empty", action="store_true",
                        help="保留空段落（默认过滤掉纯空白段落）")
    args = parser.parse_args()

    out = args.output or (os.path.splitext(args.input)[0] + "_texts.json")

    print(f"[提取] {args.input} …", flush=True)
    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    records = extract_texts(data)

    if not args.keep_empty:
        before = len(records)
        records = [r for r in records if r["content"].strip()]
        print(f"       过滤空段落：{before} → {len(records)} 条")

    with open(out, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    size_kb = os.path.getsize(out) / 1024
    print(f"[完成] 输出 → {out}  ({size_kb:.1f} KB)，共 {len(records)} 条文本")


if __name__ == "__main__":
    main()
