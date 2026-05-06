"""
docx_tools.py — Document Building Tools for DOCX Engine V3

每个公开函数将一个或多个 body_element dict 追加到模块级列表 _DOCUMENT。
输出格式完全对应 data/extraction.json 的 body_elements 结构，
可直接被 docx_compiler.py 消费。

用法示例
────────
    import docx_tools as dt

    dt.insert_abstract_with_keywords(...)
    dt.add_heading("第1章 绪论", level=1)
    dt.add_paragraph("正文内容…")
    dt.save_document("data/user_document.json")

样式参考
────────
    标题一 → OOXML style "2"  (18pt, 居中)
    标题二 → OOXML style "3"  (15pt)
    标题三 → OOXML style "4"  (14pt)
    标题四 → OOXML style "5"
    正文   → style null, firstLine=498, line=300
    摘要体 → style null, firstLine=498, line=300
    参考文献 → OOXML style "9", firstLine=498, line=300
    图/表题注 → style null, 居中, 10.5pt(sz=21)
"""

from __future__ import annotations

import base64
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional, Union

# ── OOXML namespace ────────────────────────────────────────────────────────────
W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS0 = f'xmlns:ns0="{W}"'

# ── Heading level → OOXML style id ────────────────────────────────────────────
_HEADING_STYLE: dict[int, str] = {1: "2", 2: "3", 3: "4", 4: "5"}

# ── Module-level document accumulator ─────────────────────────────────────────
_DOCUMENT: list[dict] = []
_INDEX_CTR: int = 0


# ── Public document management API ────────────────────────────────────────────

def get_document() -> list[dict]:
    """Return a shallow copy of the accumulated body_elements list."""
    return list(_DOCUMENT)


def clear_document() -> None:
    """Reset the document list and index counter."""
    global _DOCUMENT, _INDEX_CTR
    _DOCUMENT = []
    _INDEX_CTR = 0


def save_document(path: str = "data/user_document.json") -> None:
    """
    Serialize the accumulated body_elements to JSON.
    The output wraps the list in {"body_elements": [...]} so it mirrors the
    top-level shape of extraction.json and can be loaded by docx_compiler.py.
    """
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"body_elements": _DOCUMENT}, f, ensure_ascii=False, indent=2)


# ── Internal index helper ─────────────────────────────────────────────────────

def _next_idx() -> int:
    global _INDEX_CTR
    idx = _INDEX_CTR
    _INDEX_CTR += 1
    return idx


# ── XML string builders (ns0: prefix matches extraction.json serialisation) ───

def _rpr(*inner: str) -> str:
    """Wrap inner XML fragments in a w:rPr element string."""
    return f'<ns0:rPr {_NS0}>{"".join(inner)}</ns0:rPr>'


def _ppr(*inner: str) -> str:
    """Wrap inner XML fragments in a w:pPr element string."""
    return f'<ns0:pPr {_NS0}>{"".join(inner)}</ns0:pPr>'


def _pstyle(val: str) -> str:
    return f'<ns0:pStyle ns0:val="{val}" />'


def _spacing(
    *,
    before:   int | None = None,
    after:    int | None = None,
    line:     int | None = None,
    lineRule: str | None = None,
) -> str:
    attrs: list[str] = []
    if before   is not None: attrs.append(f'ns0:before="{before}"')
    if after    is not None: attrs.append(f'ns0:after="{after}"')
    if line     is not None: attrs.append(f'ns0:line="{line}"')
    if lineRule is not None: attrs.append(f'ns0:lineRule="{lineRule}"')
    return f'<ns0:spacing {" ".join(attrs)} />'


def _ind(
    *,
    firstLine:      int | None = None,
    firstLineChars: int | None = None,
    left:           int | None = None,
    leftChars:      int | None = None,
    hanging:        int | None = None,
) -> str:
    attrs: list[str] = []
    if firstLine      is not None: attrs.append(f'ns0:firstLine="{firstLine}"')
    if firstLineChars is not None: attrs.append(f'ns0:firstLineChars="{firstLineChars}"')
    if left           is not None: attrs.append(f'ns0:left="{left}"')
    if leftChars      is not None: attrs.append(f'ns0:leftChars="{leftChars}"')
    if hanging        is not None: attrs.append(f'ns0:hanging="{hanging}"')
    return f'<ns0:ind {" ".join(attrs)} />'


def _jc(val: str) -> str:
    return f'<ns0:jc ns0:val="{val}" />'


def _rfonts(
    *,
    hint:     str | None = None,
    ascii_:   str | None = None,
    hAnsi:    str | None = None,
    eastAsia: str | None = None,
) -> str:
    attrs: list[str] = []
    if hint:     attrs.append(f'ns0:hint="{hint}"')
    if ascii_:   attrs.append(f'ns0:ascii="{ascii_}"')
    if hAnsi:    attrs.append(f'ns0:hAnsi="{hAnsi}"')
    if eastAsia: attrs.append(f'ns0:eastAsia="{eastAsia}"')
    return f'<ns0:rFonts {" ".join(attrs)} />'


def _vert_align(val: str) -> str:
    return f'<ns0:vertAlign ns0:val="{val}" />'


def _sz(val: int) -> str:
    return f'<ns0:sz ns0:val="{val}" /><ns0:szCs ns0:val="{val}" />'


# ── Unified style loading ─────────────────────────────────────────────────────
# Style name mapping: internal role key → style_name in unified_style.json.
# To switch to a different style set, update the values in this dict only.
_STYLE_MAP: dict[str, str] = {
    "body":    "正文",
    "h1":      "标题一",
    "h2":      "标题二",
    "h3":      "标题三",
    "h4":      "标题四",
    "caption": "图表标题",
    "ref":     "参考文献",
}


def _load_unified_styles() -> dict[str, dict]:
    """Read unified_style.json and return fingerprints keyed by style_name."""
    path = Path(__file__).parent.parent / "data" / "unified_style.json"
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        result: dict[str, dict] = {}
        for style in data.get("standard_styles", {}).values():
            name = style.get("style_name")
            if name:
                result[name] = style.get("fingerprint", {})
        return result
    except Exception as exc:
        print(f"[WARN] unified_style.json unavailable: {exc}")
        return {}


# Fingerprints keyed by style_name; populated once at import time.
_STYLES: dict[str, dict] = _load_unified_styles()


def _style_layout(role: str) -> dict:
    """Return the layout sub-dict for the given role key (see _STYLE_MAP)."""
    return _STYLES.get(_STYLE_MAP.get(role, ""), {}).get("layout", {})


def _style_font(role: str) -> dict:
    """Return the font sub-dict for the given role key (see _STYLE_MAP)."""
    return _STYLES.get(_STYLE_MAP.get(role, ""), {}).get("font", {})


def _make_heading_ppr(role: str, pstyle_val: str, include_ea_hint: bool = False) -> str:
    """Build pPr XML for a heading level using the unified_style fingerprint.

    Conversion rule: layout.line × 2 = OOXML w:line (unified uses half the OOXML unit).
    layout.before / after are used as-is (already in OOXML twips).
    layout.jc is forwarded directly to w:jc.
    """
    lp = _style_layout(role)
    parts: list[str] = [_pstyle(pstyle_val)]
    sp: dict = {}
    if "before" in lp:
        sp["before"] = lp["before"]
    if "after" in lp:
        sp["after"] = lp["after"]
    if "line" in lp:
        sp["line"] = lp["line"] * 2
        sp["lineRule"] = "auto"
    if sp:
        parts.append(_spacing(**sp))
    if "jc" in lp:
        parts.append(_jc(lp["jc"]))
    if include_ea_hint:
        parts.append(_rpr(_EA_HINT))
    return _ppr(*parts)


# ── Reusable XML fragments ────────────────────────────────────────────────────

_SNAP_GRID  = '<ns0:snapToGrid ns0:val="0" />'
_ADJ_RIGHT  = '<ns0:adjustRightInd ns0:val="0" />'
_EA_HINT    = '<ns0:rFonts ns0:hint="eastAsia" />'  # eastAsia font hint

# run rPr for normal body runs (eastAsia font hint only)
_BODY_RPR   = _rpr(_EA_HINT)

# run rPr for superscript citation runs
_CITE_RPR   = _rpr('<ns0:kern ns0:val="0" />', _vert_align("superscript"))

# pPr for normal body / abstract paragraphs (正文 style)
_body_l = _style_layout("body")
_BODY_PPR = _ppr(
    _SNAP_GRID,
    _spacing(line=_body_l.get("line", 150) * 2, lineRule="auto"),
    _ind(firstLine=_body_l.get("firstLine", 498), firstLineChars=200),
    _rpr(_EA_HINT),
)

# pPr for keyword line (no firstLine indent)
# NOTE: hardcoded — no corresponding entry in unified_style.json; update separately
_KW_PPR = _ppr(
    _SNAP_GRID,
    _spacing(line=300, lineRule="auto"),
    _rpr(_EA_HINT),
)

# pPr templates for heading levels 1–4 (标题一~四 styles)
_HEADING_PPR: dict[int, str] = {
    1: _make_heading_ppr("h1", "2", include_ea_hint=True),
    2: _make_heading_ppr("h2", "3"),
    3: _make_heading_ppr("h3", "4"),
    4: _make_heading_ppr("h4", "5"),
}

# pPr for reference entries (style "9", 参考文献 style)
_ref_l      = _style_layout("ref")
_ref_indent = _ref_l.get("indent", {})
_REF_PPR = _ppr(
    _pstyle("1"),
    _ADJ_RIGHT,
    _SNAP_GRID,
    _jc("left"),
    _spacing(
        after=0,  # structural: suppress spacing between consecutive reference entries
        line=_ref_l.get("line", 150) * 2,
        lineRule="auto",
    ),
    _ind(
        left=_ref_indent.get("left", 543),
        hanging=_ref_indent.get("hanging", 542),
    ),
)

# pPr and run rPr for figure/table captions (图表标题 style)
_cap_l = _style_layout("caption")
_cap_f = _style_font("caption")
_CAPTION_PPR = _ppr(
    _ADJ_RIGHT,
    _jc("center"),
    _spacing(line=_cap_l.get("line", 140) * 2, lineRule="auto"),
    _rpr(_sz(_cap_f.get("size", 10) * 2)),
)
_CAPTION_RPR = _rpr(_EA_HINT, _sz(_cap_f.get("size", 10) * 2))


# ── Run / paragraph dict helpers ──────────────────────────────────────────────

def _mk_run(text: str, rpr: str | None) -> dict:
    return {"text": text, "rPr": rpr}


def _append_para(
    *,
    style:         str | None,
    text:          str,
    runs:          list[dict],
    ppr:           str | None,
    section_break: dict | None = None,
) -> dict:
    """Build a paragraph element dict and append it to _DOCUMENT."""
    elem: dict = {
        "index":         _next_idx(),
        "type":          "paragraph",
        "style":         style,
        "text":          text,
        "runs":          runs,
        "pPr":           ppr,
        "section_break": section_break,
    }
    _DOCUMENT.append(elem)
    return elem


# ── Citation-aware run splitting ──────────────────────────────────────────────

_CITE_SPLIT_RE = re.compile(r'(\[\d+(?:[,，]\d+)*\])')


def _make_runs_with_citations(text: str, base_rpr: str | None) -> list[dict]:
    """
    Split *text* on [n] / [n,m] citation tokens, producing a list of runs.
    Citation tokens are given a superscript rPr; surrounding text uses base_rpr.
    """
    parts = _CITE_SPLIT_RE.split(text)
    runs: list[dict] = []
    for part in parts:
        if not part:
            continue
        if _CITE_SPLIT_RE.fullmatch(part):
            runs.append(_mk_run(part, _CITE_RPR))
        else:
            runs.append(_mk_run(part, base_rpr))
    return runs if runs else [_mk_run(text, base_rpr)]


# ── 文本与结构 ──────────────────────────────────────────────────────────────────

def add_heading(text: str, level: int) -> None:
    """
    生成标题 body_element。

    level 1 → OOXML style "2" (黑体 18pt, 居中)
    level 2 → OOXML style "3" (黑体 15pt)
    level 3 → OOXML style "4" (黑体 14pt)
    level 4 → OOXML style "5"
    """
    style_id = _HEADING_STYLE.get(level, str(level + 1))
    ppr      = _HEADING_PPR.get(level, _HEADING_PPR[1])
    # Headings carry no explicit rPr — the named style defines the font.
    runs = [_mk_run(text, None)]
    _append_para(style=style_id, text=text, runs=runs, ppr=ppr)


def add_paragraph(text: str, style_type: str = "正文") -> None:
    """
    生成正文 body_element。

    style_type 当前支持:
        "正文"     → style null, firstLine=498, 1.5倍行距
        其他值     → 同正文, 暂无特殊映射
    文本中的 [n] 引用标注自动拆分为上标 run。
    """
    runs = _make_runs_with_citations(text, _BODY_RPR)
    _append_para(style=None, text=text, runs=runs, ppr=_BODY_PPR)


def generate_toc(max_level: int = 4) -> None:
    """
    一键生成目录占位元素。

    docx_compiler.py 遇到 type="toc" 时会生成带 w:dirty="true" 的 TOC 域代码，
    Word/WPS 首次打开文档时将自动根据标题样式刷新目录。
    """
    elem: dict = {
        "index":      _next_idx(),
        "type":       "toc",
        "max_level":  max_level,
        #"style_prefix": style_prefix,
    }
    _DOCUMENT.append(elem)


def insert_page_break() -> None:
    """
    换页: 生成一个含 w:br type="page" 的段落元素。
    """
    br_rpr = _rpr(_EA_HINT)
    br_run = {
        "text": "",
        "rPr":  br_rpr,
        "break_type": "page",  # docx_compiler 识别此标志插入 <w:br w:type="page"/>
    }
    _append_para(
        style=None,
        text="",
        runs=[br_run],
        ppr=_ppr(_SNAP_GRID, _spacing(line=300, lineRule="auto")),
    )


def insert_section_break(
    header_template: Optional[Literal["default", "even", "first", "none"]] = None,
    footer_template: Optional[Literal["default", "even", "first", "none"]] = None,
    restart_page_number: Optional[int] = None,
    header_refs: Optional[Dict[str, str]] = None,
    footer_refs: Optional[Dict[str, str]] = None,
) -> None:
    """
    换节。三个参数均描述下一节的属性。

    header_refs / footer_refs (优先):
        直接指定真实 rId 映射, 如 {"default": "rId10", "even": "rId11"}。
        与 extraction.json sections 格式一致, 由 docx_compiler 解析。

    header_template / footer_template (回退):
        当未提供 header_refs/footer_refs 时使用模板名生成占位符 refs。

    restart_page_number:
        下一节起始页码; None 表示续接。
    """
    # Explicit rId refs take priority over template-name stubs.
    if header_refs is None:
        header_refs = {}
        if header_template and header_template != "none":
            header_refs["default"] = "__template__"
            if header_template == "even":
                header_refs["even"] = "__template__"
            elif header_template == "first":
                header_refs["first"] = "__template__"

    if footer_refs is None:
        footer_refs = {}
        if footer_template and footer_template != "none":
            footer_refs["default"] = "__template__"
            if footer_template == "even":
                footer_refs["even"] = "__template__"

    section_break: dict = {
        "header_refs": header_refs,
        "footer_refs": footer_refs,
        "page_size":   {"w": "11906", "h": "16838"},  # A4
    }
    if restart_page_number is not None:
        section_break["restart_page_number"] = restart_page_number

    _append_para(
        style=None,
        text="",
        runs=[],
        ppr=_ppr(_ADJ_RIGHT, _SNAP_GRID, _rpr(_EA_HINT)),
        section_break=section_break,
    )


# ── 图片 ────────────────────────────────────────────────────────────────────────

def insert_figure(
    data_source: Union[str, bytes, None] = None,
    base64:      Optional[str]           = None,   # alias: user_data.json 直接传入 base64 字符串
    width: float = 0.0,
    height: float = 0.0,
    caption: Optional[str] = None,
    position: Union[
        Literal["left", "center", "right"],
        Dict[str, Any],
    ] = "center",
    drawing_xml: Optional[str] = None,
) -> None:
    """
    插入图片元素。

    drawing_xml (优先):
        直接提供完整的 DrawingML/AlternateContent XML 字符串
        (来自 extraction.json runs[].drawing_xml 字段)。
        docx_compiler 将其原样嵌入 <w:r> 中, 保留所有定位与关系 rId。

    data_source / base64 (回退, 二者取其一):
        base64      → user_data.json 直接传入的 base64 字符串, 优先于 data_source
        data_source → str 文件路径 或 bytes 原始数据

    position:
        字符串 "left"/"center"/"right"
        或 dict {"align": "center", "indent_left": 20, "wrap": "tight"}
    """
    # Normalise position
    if isinstance(position, dict):
        align = position.get("align", "center")
    else:
        align = position

    if drawing_xml:
        # drawing_xml path — no base64 needed; compiler re-emits directly.
        elem: dict = {
            "index":       _next_idx(),
            "type":        "image",
            "drawing_xml": drawing_xml,
            "caption":     caption,
            "position":    align,
        }
        _DOCUMENT.append(elem)
    else:
        # Resolve the b64 string.
        # Import the module under an alias to avoid shadowing by the `base64` parameter.
        import base64 as _b64lib

        if base64:
            # Parameter is already a base64-encoded string (from user_data.json).
            b64 = base64
        elif isinstance(data_source, (str, Path)):
            try:
                raw: bytes = Path(data_source).read_bytes()
            except (OSError, FileNotFoundError):
                raw = b""
            b64 = _b64lib.b64encode(raw).decode("ascii") if raw else ""
        elif isinstance(data_source, bytes):
            b64 = _b64lib.b64encode(data_source).decode("ascii") if data_source else ""
        else:
            b64 = ""

        elem = {
            "index":    _next_idx(),
            "type":     "image",
            "base64":   b64,
            "width":    width,
            "height":   height,
            "caption":  caption,
            "position": align,
        }
        _DOCUMENT.append(elem)

    if caption:
        _append_para(
            style=None,
            text=caption,
            runs=[_mk_run(caption, _CAPTION_RPR)],
            ppr=_CAPTION_PPR,
        )


# ── 表格 ────────────────────────────────────────────────────────────────────────

def insert_table(
    rows: List[List[Any]],
    caption: Optional[str] = None,
    auto_format: bool = True,
    column_widths: Optional[List[float]] = None,
) -> None:
    """
    插入表格 body_element。

    data 第一行视为表头。
    每个单元格转为 extraction.json 的 cell 结构 (paragraphs list)。
    """
    if caption:
        # Table caption appears ABOVE the table (Chinese academic convention).
        _append_para(
            style=None,
            text=caption,
            runs=[_mk_run(caption, _CAPTION_RPR)],
            ppr=_CAPTION_PPR,
        )

    # Build rows in extraction.json table format
    built_rows: list[list[dict]] = []
    for row_data in rows:
        row: list[dict] = []
        for cell_val in row_data:
            cell_text = str(cell_val) if not isinstance(cell_val, str) else cell_val
            cell_rpr  = _rpr(_EA_HINT, _sz(21))
            cell_ppr  = _ppr(
                _ADJ_RIGHT,
                _jc("center"),
                _rpr(_EA_HINT, _sz(21)),
            )
            cell: dict = {
                "text": cell_text,
                "paragraphs": [
                    {
                        "index":         0,
                        "type":          "paragraph",
                        "style":         None,
                        "text":          cell_text,
                        "runs":          [_mk_run(cell_text, cell_rpr)],
                        "pPr":           cell_ppr,
                        "section_break": None,
                    }
                ],
            }
            row.append(cell)
        built_rows.append(row)

    elem: dict = {
        "index":         _next_idx(),
        "type":          "table",
        "rows":          built_rows,
        "auto_format":   auto_format,
        "column_widths": column_widths,
    }
    _DOCUMENT.append(elem)


# ── 公式 ────────────────────────────────────────────────────────────────────────

def insert_equation(
    expression:   str                              = "",
    category:     Optional[Literal["omath", "ole"]] = None,
    position:     Literal["left", "center", "right"] = "center",
    suffix:       Optional[str]   = None,   # 如 "(4-1)", "(5-2)"
    suffix_position: Literal["right", "new_line"] = "right",
    # Structured fields produced by user_data_generator (all optional):
    omml:         Optional[str]   = None,   # OMML XML fragment → omath 模式
    ole_base64:   Optional[str]   = None,   # OLE 二进制 base64 → ole 模式
    image_base64: Optional[str]   = None,   # OLE 预览图 base64（可选，暂存备用）
    label:        Optional[str]   = None,   # 公式编号，如 "(1-1)"；OLE 无法渲染时作占位文字
    width_pt:     Optional[float] = None,   # OLE 显示宽度（点）
    height_pt:    Optional[float] = None,   # OLE 显示高度（点）
    prog_id:      Optional[str]   = None,   # OLE ProgID（默认 Equation.3）
    text_before:  Optional[str]   = None,   # 公式前的行内文字，如 "式中"
    text_after:   Optional[str]   = None,   # 公式后的行内文字，如 "——粒子直径（m）；"
    is_inline:    Optional[bool]  = None,   # True → 行内公式（嵌入正文句子中）
) -> None:
    """
    插入公式 body_element。

    模式自动推断规则（优先级从高到低）：
      1. ole_base64 非空  → category="ole"，base64 数据传入编译器写入 word/embeddings/
      2. omml 非空        → category="omath"，OMML XML 作为 formula 字段
      3. 以上均空         → 沿用 category 参数（默认 "omath"），formula 取 expression

    suffix          : 公式编号, 如 "(4-1)"
    suffix_position : "right"    → 编号紧跟公式同行
                      "new_line" → 编号另起一行
    text_before / text_after:
        行内公式时公式前后的文字片段，由 docx_compiler 拼入同一段落。
    is_inline:
        True 时编译器将公式嵌入文本段落；False/None 时按独立居中公式行排版。
    """
    # ── 自动推断 category ────────────────────────────────────────────────────────
    if ole_base64:
        category = "ole"
    elif omml:
        category = "omath"
    elif category is None:
        category = "omath"

    # ── 解析 formula 字段 ────────────────────────────────────────────────────────
    if category == "omath":
        formula = omml or expression
    else:  # ole
        formula = label or expression

    # label 同时作为公式编号 suffix 的来源（当 suffix 未单独传入时）
    resolved_suffix = suffix or label or ""

    elem: dict = {
        "index":           _next_idx(),
        "type":            category,
        "formula":         formula,
        "formula_index":   resolved_suffix,
        "position":        position,
        "suffix_position": suffix_position,
    }
    if is_inline:
        elem["is_inline"] = True
    if text_before:
        elem["text_before"] = text_before
    if text_after:
        elem["text_after"] = text_after
    if category == "ole":
        elem["base64"] = ole_base64 or ""
        elem["image_base64"] = image_base64 or ""
        if width_pt  is not None: elem["width_pt"]  = width_pt
        if height_pt is not None: elem["height_pt"] = height_pt
        if prog_id   is not None: elem["prog_id"]   = prog_id
    _DOCUMENT.append(elem)


# ── 摘要与关键词 ──────────────────────────────────────────────────────────────────

def insert_abstract_with_keywords(
    cn_content: str,
    en_content: str,
    cn_keywords: List[str],
    en_keywords: List[str],
    cn_title: str = "摘  要",
    en_title: str = "Abstract",
    keyword_label_cn: str = "关键词",
    keyword_label_en: str = "Keywords",
    cn_section_break: Optional[Dict] = None,
    en_section_break: Optional[Dict] = None,
) -> None:
    """
    插入摘要及关键词, 生成以下段落序列:

        [标题一] cn_title
        [正文  ] cn_content  (style null, firstLine indent)
        [关键词] 关键词:kw1；kw2；…  (label 黑体, 关键词宋体)  ← cn_section_break 嵌入此段
        [标题一] en_title
        [正文  ] en_content
        [关键词] Keywords：kw1, kw2, …  (Keywords 加粗)        ← en_section_break 嵌入此段

    cn_section_break / en_section_break:
        分节信息 dict，格式与 _append_para section_break 参数一致：
        {"header_refs": {...}, "footer_refs": None, "page_size": {...}}
        直接嵌入对应关键词段落的 section_break 字段，绝对不会产生空白页。
    """
    # ── Chinese abstract ─────────────────────────────────────────────────────
    add_heading(cn_title, level=1)

    cn_runs = _make_runs_with_citations(cn_content, _BODY_RPR)
    _append_para(style=None, text=cn_content, runs=cn_runs, ppr=_BODY_PPR)

    # 关键词: label in 黑体, keywords in 宋体 separated by ；
    cn_kw_label_rpr = _rpr(
        _rfonts(hint="eastAsia", ascii_="黑体", hAnsi="黑体", eastAsia="黑体")
    )
    cn_kw_body_rpr = _rpr(_EA_HINT)
    cn_kw_text     = f"{keyword_label_cn}:{'；'.join(cn_keywords)}"
    cn_kw_runs     = [
        _mk_run(keyword_label_cn, cn_kw_label_rpr),
        _mk_run(f":{'；'.join(cn_keywords)}", cn_kw_body_rpr),
    ]
    _append_para(style=None, text=cn_kw_text, runs=cn_kw_runs, ppr=_KW_PPR,
                 section_break=cn_section_break)

    # ── English abstract ─────────────────────────────────────────────────────
    add_heading(en_title, level=1)

    en_runs = _make_runs_with_citations(en_content, _BODY_RPR)
    _append_para(style=None, text=en_content, runs=en_runs, ppr=_BODY_PPR)

    # Keywords：label in bold, keywords plain, separated by ", "
    en_kw_label_rpr = _rpr('<ns0:b /><ns0:bCs />')
    en_kw_colon_rpr = _rpr(_EA_HINT)
    en_kw_body_rpr  = None  # no rPr — plain English
    en_kw_text      = f"{keyword_label_en}：{', '.join(en_keywords)}"
    en_kw_runs      = [
        _mk_run(keyword_label_en, en_kw_label_rpr),
        _mk_run("：",             en_kw_colon_rpr),
        _mk_run(", ".join(en_keywords), en_kw_body_rpr),
    ]
    _append_para(style=None, text=en_kw_text, runs=en_kw_runs, ppr=_KW_PPR,
                 section_break=en_section_break)


# ── 参考文献 ──────────────────────────────────────────────────────────────────────

def add_reference(
    index: int,
    content: str,
    before: Optional[str] = None,
    after: Optional[str] = None,
    auto_cite: bool = True,
) -> None:
    """
    生成参考文献条目段落 (style "9") 并将行内引用标注写入元数据。

    before / after:
        触发该文献的正文上下文, 供 docx_compiler 或后处理工具将行内
        "[index]" 上标插入对应正文段落。
        auto_cite=True 时, 标志 docx_compiler 自动扫描正文并插入上标。

    参考文献文本以 "[index] content" 格式写入, 符合 GB/T 7714 编排惯例。
    """
    full_text = f"［{index}］{content}"
    run_rpr   = _rpr(_EA_HINT)
    runs      = [_mk_run(full_text, run_rpr)]

    elem = _append_para(style="1", text=full_text, runs=runs, ppr=_REF_PPR)

    # Attach citation context as extra metadata fields (not in base extraction
    # schema, but harmless — compiler ignores unknown keys).
    elem["citation_index"] = index
    elem["citation_before"] = before
    elem["citation_after"]  = after
    elem["auto_cite"]       = auto_cite
