#!/usr/bin/env python3
"""
docx_inline_tag_engine.py  v7
==============================
模板 + 数据分离架构

v5 新增：目录（TOC）支持
  ① JSON 里加 {"type": "toc"} 块，引擎在该位置生成目录
  ② 双轨渲染模式：
       auto  — 插入 Word TOC 域（打开时自动刷新，格式由 TOC 1/2/3 样式决定）
       manual— 用提取的模板样式手动排版（页码由 JSON 提供，适合精确控制）
  ③ toc_exclude：在 section / heading 块上设为 true，该标题不进入目录
  ④ 模板 [[目录]] 块：放三行示例条目（1/2/3级格式），引擎提取其 pPr/rPr

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

JSON 完整结构（v5）
-------------------
{
  "page_footer_spec": "...",           // 自然语言页脚描述，LLM 解析
  // 或者直接:
  "page_footer_config": [...],

  "toc_mode": "auto",                  // "auto"（默认）或 "manual"
  // manual 模式下需要提供每个标题的页码：
  // "toc_entries": [
  //   {"title": "摘  要",   "level": 1, "page": "i"},
  //   {"title": "第一章 引言", "level": 1, "page": 1},
  //   ...
  // ]

  "content": [
    // ── 前置章节（toc_exclude 控制是否进目录）──────────────────
    {"type": "section", "section_type": "abstract",
     "toc_exclude": true,              // 不进目录（默认 false）
     "value": "本文研究了…"},

    // ── 目录占位（放在想要目录出现的位置）─────────────────────
    {"type": "toc",
     "title": "目  录",                // 可选，目录章节标题
     "toc_title_exclude": true},       // 目录标题本身是否排除出目录（默认 true）

    // ── 正文 ────────────────────────────────────────────────────
    {"type": "heading1", "value": "第一章 引言"},
    {"type": "heading2", "value": "1.1 研究背景"},
    {"type": "body",     "value": "正文…"},
    {"type": "table",    "caption": "表1", "data": [[...]]},
    {"type": "formula",  "label": "式(1)", "latex": "..."},
    {"type": "image",    "path": "fig.png", "caption": "图1"},

    // ── 后置章节 ────────────────────────────────────────────────
    {"type": "section", "section_type": "conclusion",
     "toc_exclude": false,             // 进目录
     "value": "综上所述…"},

    {"type": "section", "section_type": "acknowledgement",
     "toc_exclude": true,
     "value": "感谢…"},

    {"type": "section", "section_type": "publications",
     "toc_exclude": false,
     "value": "1. …"}
  ],

  "references": [...],
  "citations":  [...]
}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

模板 [[目录]] 块写法（决定手动目录的格式，auto 模式也可以有）
--------------------------------------------------------------
[[目录]]
第一章 示例一级条目……1        ← 设好字体/加粗/缩进/前导符，这一行的格式 → TOC level 1
1.1 示例二级条目……1           ← 这一行的格式 → TOC level 2
1.1.1 示例三级条目……1         ← 这一行的格式 → TOC level 3
[[目录]]

三行各自的 pPr（缩进/tab/间距）和 rPr（字体/字号/加粗）会被分别提取。
如果只放两行，三级目录会复用二级样式；只放一行则全部复用该样式。
auto 模式下提取结果用于向 Word 样式表写入 TOC 1/2/3（如果不写，Word 会用自己的默认样式）。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

依赖：
  pip install python-docx --break-system-packages
  pip install latex2mathml --break-system-packages   # formula/latex 可选
"""

import sys
from docx_fixer.api import process


def main():
    if len(sys.argv) < 4:
        print("用法: python main.py <模板.docx> <数据.json> <输出.docx> [api_key],使用默认参数")
        sys.argv.extend(["data\\full_template_v6.docx","data\\full_user_data.json","data\\output.docx"])
    

    template_path = sys.argv[1] 
    data_path = sys.argv[2]
    output_path = sys.argv[3]
    api_key = sys.argv[4] if len(sys.argv) > 4 else None

    process(template_path, data_path, output_path, api_key)


if __name__ == "__main__":
    main()
