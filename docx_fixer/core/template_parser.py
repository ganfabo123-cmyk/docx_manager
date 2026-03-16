from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import copy
from typing import Optional
from .constants import W, _STATIC_TAGS, TYPE_TO_TAG, TAG_RE
from ..models.models import StyleTemplate, HeadingStyle, TocLevelStyle


def _resolve_para_style(doc, p_elem):
    """
    从段落元素提取完整的 pPr 和 rPr，解析样式继承链。

    Word 格式优先级（从高到低）：
      直接段落格式(pPr) > 段落样式(pStyle) > 默认段落样式

    返回 (merged_pPr, merged_rPr)：
      merged_pPr: 直接pPr merge 样式pPr（直接格式优先）
      merged_rPr: 直接rPr merge 样式rPr（直接格式优先）
    """
    direct_pPr = p_elem.find(f"{{{W}}}pPr")
    
    # 找到 pStyle
    pStyle_val = None
    if direct_pPr is not None:
        pStyle_elem = direct_pPr.find(f"{{{W}}}pStyle")
        if pStyle_elem is not None:
            pStyle_val = pStyle_elem.get(qn("w:val"))

    # 从样式表找样式定义
    style_pPr = None
    style_rPr = None
    if pStyle_val is not None:
        try:
            styles_elem = doc.part.styles.element
            style_def = styles_elem.find(
                f".//{{{W}}}style[@{{{W}}}styleId='{pStyle_val}']"
            )
            if style_def is not None:
                style_pPr = style_def.find(f"{{{W}}}pPr")
                style_rPr = style_def.find(f"{{{W}}}rPr")
        except Exception:
            pass

    # 直接 rPr（在第一个 run 里，或段落 pPr 里的 rPr）
    direct_rPr = None
    # 先找段落级 rPr（pPr 里的 rPr，用于段落标记字符格式）
    if direct_pPr is not None:
        direct_rPr = direct_pPr.find(f"{{{W}}}rPr")
    # 再找第一个 run 的 rPr（优先级更高，代表实际文字格式）
    first_run = p_elem.find(f"{{{W}}}r")
    if first_run is not None:
        run_rPr = first_run.find(f"{{{W}}}rPr")
        if run_rPr is not None:
            direct_rPr = run_rPr   # run 的直接 rPr 优先

    # 合并 pPr：样式pPr 为底，直接格式覆盖
    merged_pPr = OxmlElement("w:pPr")
    # 先写样式 pPr（排除 pStyle 本身）
    if style_pPr is not None:
        for child in style_pPr:
            tag_local = child.tag.split("}")[-1]
            if tag_local != "pStyle":
                if merged_pPr.find(child.tag) is None:
                    merged_pPr.append(copy.deepcopy(child))
    # 再用直接 pPr 覆盖（排除 pStyle）
    if direct_pPr is not None:
        for child in direct_pPr:
            tag_local = child.tag.split("}")[-1]
            if tag_local in ("pStyle", "rPr"):
                continue
            existing = merged_pPr.find(child.tag)
            if existing is not None:
                merged_pPr.remove(existing)
            merged_pPr.append(copy.deepcopy(child))

    # 合并 rPr：样式rPr 为底，直接格式覆盖
    merged_rPr = OxmlElement("w:rPr")
    if style_rPr is not None:
        for child in style_rPr:
            if merged_rPr.find(child.tag) is None:
                merged_rPr.append(copy.deepcopy(child))
    if direct_rPr is not None:
        for child in direct_rPr:
            existing = merged_rPr.find(child.tag)
            if existing is not None:
                merged_rPr.remove(existing)
            merged_rPr.append(copy.deepcopy(child))

    # 若合并结果为空，返回 None
    pPr_result = merged_pPr if len(merged_pPr) > 0 else None
    rPr_result = merged_rPr if len(merged_rPr) > 0 else None
    return pPr_result, rPr_result


def parse_template(doc_path: str, extra_section_types: list[str] | None = None) -> StyleTemplate:
    doc  = Document(doc_path)
    body = doc.element.body
    st   = StyleTemplate()

    section_tags       = set(extra_section_types or [])
    section_title_tags = {f"{s}标题" for s in section_tags}
    all_tags = _STATIC_TAGS | section_tags | section_title_tags

    in_block: Optional[str] = None
    buf: list = []

    for child in body:
        local = child.tag.split("}")[-1]

        if local == "p":
            text = _para_text(child).strip()
            m    = TAG_RE.match(text)
            tag  = m.group(1) if m else None

            if tag in all_tags and in_block is None:
                in_block = tag
                buf = []
                continue

            if tag == in_block and in_block is not None:
                _extract_style(st, in_block, buf, section_tags, doc)
                in_block = None
                buf = []
                continue

            if in_block is not None:
                buf.append(child)

        elif local in ("tbl", "oMathPara"):
            if in_block is not None:
                buf.append(child)

    return st


def _extract_style(st: StyleTemplate, tag: str, elems: list, section_tags: set[str], doc=None):

    if tag == "表格":
        for e in elems:
            local = e.tag.split("}")[-1]
            if local == "p" and st.table_caption_pPr is None:
                pPr, rPr = _resolve_para_style(doc, e)
                st.table_caption_pPr = pPr
                st.table_caption_rPr = rPr
            elif local == "tbl" and st.table_proto is None:
                st.table_proto = copy.deepcopy(e)

    elif tag in ("一级标题", "二级标题", "三级标题"):
        key = {v: k for k, v in TYPE_TO_TAG.items()}[tag]
        for e in elems:
            if e.tag.split("}")[-1] == "p":
                pPr, rPr = _resolve_para_style(doc, e)
                st.heading_styles[key] = HeadingStyle(pPr=pPr, rPr=rPr)
                break

    elif tag == "正文":
        for e in elems:
            if e.tag.split("}")[-1] == "p":
                pPr, rPr = _resolve_para_style(doc, e)
                st.body_pPr = pPr
                st.body_rPr = rPr
                break

    elif tag == "参考文献":
        for e in elems:
            if e.tag.split("}")[-1] == "p":
                pPr, rPr = _resolve_para_style(doc, e)
                st.ref_pPr_proto = pPr
                st.ref_rPr_proto = rPr
                break

    elif tag == "公式":
        paras = [e for e in elems if e.tag.split("}")[-1] == "p"]
        if paras:
            pPr, _ = _resolve_para_style(doc, paras[0])
            st.formula_pPr = pPr
        if len(paras) >= 2:
            _, rPr = _resolve_para_style(doc, paras[1])
            st.formula_label_rPr = rPr

    elif tag == "图片":
        # 约定：块内第一行 → 图片段落格式，第二行 → caption 格式
        paras = [e for e in elems if e.tag.split("}")[-1] == "p"]
        if len(paras) >= 1:
            pPr, _ = _resolve_para_style(doc, paras[0])
            st.image_pPr = pPr
        if len(paras) >= 2:
            pPr2, rPr2 = _resolve_para_style(doc, paras[1])
            st.image_caption_pPr = pPr2
            st.caption_rPr = rPr2

    elif tag == "目录":
        # 提取 1/2/3 级目录条目样式
        # 块内每一个 <w:p> 对应一个级别，按顺序收集
        para_elems = [e for e in elems if e.tag.split("}")[-1] == "p"]
        for pe in para_elems[:3]:   # 最多取前三行
            pPr = pe.find(f"{{{W}}}pPr")
            rPr = pe.find(f".//{{{W}}}rPr")
            st.toc_level_styles.append(TocLevelStyle(
                pPr = copy.deepcopy(pPr) if pPr is not None else None,
                rPr = copy.deepcopy(rPr) if rPr is not None else None,
            ))

    elif tag in section_tags:
        ss = st.section_styles.setdefault(tag, {})
        for e in elems:
            if e.tag.split("}")[-1] == "p":
                pPr, rPr = _resolve_para_style(doc, e)
                if pPr is not None:
                    ss["body_pPr"] = pPr
                if rPr is not None:
                    ss["body_rPr"] = rPr
                break

    elif tag.endswith("标题") and tag[:-2] in section_tags:
        stype = tag[:-2]
        ss = st.section_styles.setdefault(stype, {})
        for e in elems:
            if e.tag.split("}")[-1] == "p":
                pPr, rPr = _resolve_para_style(doc, e)
                if pPr is not None:
                    ss["title_pPr"] = pPr
                if rPr is not None:
                    ss["title_rPr"] = rPr
                break


def _para_text(p_elem) -> str:
    return "".join(t.text or "" for t in p_elem.iter(f"{{{W}}}t"))
