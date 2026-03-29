from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import copy

from .constants import W, M


def _render_formula(doc, item, st):
    label = item.get("label", "")
    if "omml" in item:
        omml_elem = _parse_omml_string(item["omml"])
    elif "latex" in item:
        omml_elem = _latex_to_omml(item["latex"])
    else:
        omml_elem = None

    body   = doc.element.body
    sectPr = body.find(f"{{{W}}}sectPr")

    if omml_elem is not None:
        p_elem = (_build_formula_para_with_label(omml_elem, label, st)
                  if label else _build_formula_para(omml_elem, st))
        if sectPr is not None:
            sectPr.addprevious(p_elem)
        else:
            body.append(p_elem)
    else:
        raw  = item.get("latex", item.get("omml", ""))
        txt  = f"{raw}    {label}" if label else raw
        # 降级为纯文本，格式完全来自模板 formula_pPr，不加硬编码对齐
        para = doc.add_paragraph()
        _apply_pPr(para._p, st.formula_pPr)
        para.add_run(txt)


def _build_formula_para(omml_elem, st):
    p = OxmlElement("w:p")
    p.append(_make_formula_pPr(st))
    p.append(copy.deepcopy(omml_elem))
    return p


def _build_formula_para_with_label(omml_elem, label, st):
    """
    带编号的公式行：[Tab] [公式居中] [Tab] [式(x)右对齐]
    tab stop 位置从模板 formula_pPr 里提取；若无，则从模板正文宽度推断。
    label rPr 使用模板 formula_label_rPr（[[公式]] 块第二行）。
    """
    p = OxmlElement("w:p")
    pPr = _make_formula_pPr(st)

    # 确保 pPr 里有公式行所需的两个 tab stop
    # 优先用模板 formula_pPr 里已有的 tabs；若无则用推断值（中心/右端）
    existing_tabs = pPr.find(f"{{{W}}}tabs")
    if existing_tabs is None:
        tabs = OxmlElement("w:tabs")
        for val, pos in [("center", "4680"), ("right", "9360")]:
            tab = OxmlElement("w:tab")
            tab.set(qn("w:val"), val)
            tab.set(qn("w:pos"), pos)
            tabs.append(tab)
        pPr.append(tabs)

    p.append(pPr)

    def tab_run():
        r = OxmlElement("w:r")
        r.append(OxmlElement("w:tab"))
        return r

    p.append(tab_run())
    p.append(copy.deepcopy(omml_elem))
    p.append(tab_run())

    r_label = OxmlElement("w:r")
    if st.formula_label_rPr is not None:
        r_label.append(copy.deepcopy(st.formula_label_rPr))
    t_label = OxmlElement("w:t")
    t_label.text = label
    r_label.append(t_label)
    p.append(r_label)
    return p


def _make_formula_pPr(st):
    """公式段落 pPr：完全来自模板 [[公式]] 块，无模板则返回空 pPr（不加任何硬编码格式）"""
    if st.formula_pPr is not None:
        return copy.deepcopy(st.formula_pPr)
    return OxmlElement("w:pPr")


def _parse_omml_string(omml_str):
    from lxml import etree
    try:
        return etree.fromstring(omml_str.encode())
    except Exception as exc:
        raise ValueError(f"无效的 OMML XML：{exc}") from exc


def _latex_to_omml(latex):
    try:
        import latex2mathml.converter
        from lxml import etree
    except ImportError:
        return None
    mathml_str  = latex2mathml.converter.convert(latex)
    mathml_elem = etree.fromstring(mathml_str.encode())
    return _mathml_to_omml_via_lxml(mathml_elem)


def _mathml_to_omml_via_lxml(mathml_elem):
    def _tag(e):
        return e.tag.split("}")[-1] if "}" in e.tag else e.tag

    def _text_run(text):
        r = OxmlElement("m:r")
        t = OxmlElement("m:t")
        t.text = text
        r.append(t)
        return r

    def _append_c(parent, elem):
        result = _conv(elem)
        if result is None: return
        if isinstance(result, list):
            for it in result:
                if it is not None: parent.append(it)
        else:
            parent.append(result)

    def _conv(elem):
        tag      = _tag(elem)
        children = list(elem)
        if tag in ("math", "mrow", "mstyle", "mpadded", "mphantom"):
            container = OxmlElement("m:oMath") if tag == "math" else None
            results   = [_conv(c) for c in children]
            if container is not None:
                for r in results:
                    if r is None: continue
                    (container.extend(r) if isinstance(r, list)
                     else container.append(r))
                return container
            return results
        elif tag == "mfrac":
            f = OxmlElement("m:f"); f.append(OxmlElement("m:fPr"))
            n = OxmlElement("m:num"); d = OxmlElement("m:den")
            if children: _append_c(n, children[0])
            if len(children)>1: _append_c(d, children[1])
            f.append(n); f.append(d); return f
        elif tag == "msup":
            ss = OxmlElement("m:sSup")
            e_ = OxmlElement("m:e"); s_ = OxmlElement("m:sup")
            if children: _append_c(e_, children[0])
            if len(children)>1: _append_c(s_, children[1])
            ss.append(e_); ss.append(s_); return ss
        elif tag == "msub":
            ss = OxmlElement("m:sSub")
            e_ = OxmlElement("m:e"); s_ = OxmlElement("m:sub")
            if children: _append_c(e_, children[0])
            if len(children)>1: _append_c(s_, children[1])
            ss.append(e_); ss.append(s_); return ss
        elif tag == "msqrt":
            rad = OxmlElement("m:rad")
            rp = OxmlElement("m:radPr"); dh = OxmlElement("m:degHide")
            dh.set(qn("m:val"), "1"); rp.append(dh); rad.append(rp)
            rad.append(OxmlElement("m:deg"))
            e_ = OxmlElement("m:e")
            for c in children: _append_c(e_, c)
            rad.append(e_); return rad
        elif tag in ("mn", "mi", "mo", "mtext", "ms"):
            text = (elem.text or "").strip()
            return _text_run(text) if text else None
        else:
            parts = [_conv(c) for c in children]
            parts = [p for p in parts if p is not None]
            if not parts and elem.text:
                return _text_run(elem.text.strip())
            if len(parts) == 1 and not isinstance(parts[0], list):
                return parts[0]
            wrap = OxmlElement("m:oMath")
            for part in parts:
                if isinstance(part, list):
                    for pp in part:
                        if pp is not None: wrap.append(pp)
                else:
                    wrap.append(part)
            return wrap

    oMath = _conv(mathml_elem)
    if oMath is None:
        oMath = OxmlElement("m:oMath")
        oMath.append(_text_run("?"))
    oMathPara = OxmlElement("m:oMathPara")
    oMathPara.append(oMath)
    return oMathPara


def _apply_pPr(p_elem, pPr_proto) -> None:
    """完整替换段落 pPr（不继承）。sectPr 子元素保留（分节符不能丢）。"""
    if pPr_proto is None: return
    existing = p_elem.find(f"{{{W}}}pPr")
    # 保留 sectPr（分节符），其他全部替换
    saved_sectPr = None
    if existing is not None:
        saved_sectPr = existing.find(f"{{{W}}}sectPr")
        p_elem.remove(existing)
    new_pPr = copy.deepcopy(pPr_proto)
    # 去掉 pStyle（输出文档不继承样式）
    pStyle = new_pPr.find(f"{{{W}}}pStyle")
    if pStyle is not None:
        new_pPr.remove(pStyle)
    if saved_sectPr is not None:
        new_pPr.append(copy.deepcopy(saved_sectPr))
    p_elem.insert(0, new_pPr)
