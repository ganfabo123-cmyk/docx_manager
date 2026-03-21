import json
from pathlib import Path
from docx import Document

from ..core.template_parser import parse_template
from ..core.generate import generate
from ..core.footer import parse_footer_spec


def process(template_path: str, data_path: str, output_path: str, api_key: str | None = None, verbose: bool = True) -> Document:
    """
    一步到位：模板 + json → 输出 docx

    参数
    ----
    template_path : 含 [[标记]] 的格式模板 .docx
    data_path     : user_data.json 路径
    output_path   : 输出 .docx 路径
    api_key       : Anthropic API Key（LLM 解析页脚描述，可不填）
    """
    data = json.loads(Path(data_path).read_text(encoding="utf-8"))

    section_types = list({
        item["section_type"]
        for item in data.get("content", [])
        if item.get("type") == "section"
    })

    st = parse_template(template_path, extra_section_types=section_types)

    # 页脚配置
    footer_config = None
    if "page_footer_config" in data:
        footer_config = data["page_footer_config"]
        if verbose:
            print(f"📑  页脚配置（直接）：{footer_config}")
    elif "page_footer_spec" in data:
        if verbose:
            print(f"🤖  解析页脚描述：{data['page_footer_spec']!r}")
        footer_config = parse_footer_spec(data["page_footer_spec"], api_key)
        if verbose:
            print(f"   → 解析结果：{footer_config}")

    if verbose:
        content      = data.get("content", [])
        n_section    = sum(1 for x in content if x.get("type") == "section")
        n_toc        = sum(1 for x in content if x.get("type") == "toc")
        toc_mode     = data.get("toc_mode", "auto")
        toc_styles_n = len(st.toc_level_styles)
        print(f"📐  模板：表格={'有' if st.table_proto is not None else '无'}  "
              f"标题样式={list(st.heading_styles.keys())}  "
              f"正文pPr={'有' if st.body_pPr is not None else '无'}  "
              f"文献pPr={'有' if st.ref_pPr_proto is not None else '无'}  "
              f"章节样式={list(st.section_styles.keys()) or '无'}  "
              f"TOC级样式={toc_styles_n}级")
        print(f"📋  内容：{len(content)} 块（{n_section} 章节块，{n_toc} 目录块，" 
              f"TOC模式={toc_mode}）  "
              f"文献={len(data.get('references',[]))}  "
              f"引用={len(data.get('citations',[]))}")

    doc = generate(data, st, output_path, footer_config=footer_config)

    if verbose:
        print(f"✅  已生成：{output_path}")
    return doc
