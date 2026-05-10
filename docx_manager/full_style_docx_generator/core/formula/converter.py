import re
from pathlib import Path
import latex2mathml.converter
from lxml import etree

_DOLLAR_INLINE = re.compile(r'\$([^$\n]+?)\$')

_XSL_PATH = Path(__file__).parent / "assets" / "MML2OMML.XSL"
_transform = None

_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _get_transform():
    global _transform
    if _transform is None:
        xslt_doc = etree.parse(str(_XSL_PATH))
        _transform = etree.XSLT(xslt_doc)
    return _transform


def _fix_script_delimiters(root):
    # 2026-05-10: MML2OMML.XSL 把上下标内的 (t) 转成 <m:d>（自动伸缩括号），
    # 在 <m:sup>/<m:sub> 内 Word 无法正确渲染，导致 ^(t) 可见甚至触发 linear 模式。
    # 将 <m:d> 替换为显式括号 run，确保括号以固定大小渲染。
    def make_run(text):
        r = etree.Element(f"{{{_M_NS}}}r")
        t = etree.SubElement(r, f"{{{_M_NS}}}t")
        t.text = text
        return r

    def get_delimiter_info(d_elem):
        dPr = d_elem.find(f"{{{_M_NS}}}dPr")
        beg, end, sep = "(", ")", ""
        if dPr is not None:
            b = dPr.find(f"{{{_M_NS}}}begChr")
            e = dPr.find(f"{{{_M_NS}}}endChr")
            s = dPr.find(f"{{{_M_NS}}}sepChr")
            if b is not None:
                beg = b.get(f"{{{_M_NS}}}val", "(")
            if e is not None:
                end = e.get(f"{{{_M_NS}}}val", ")")
            if s is not None:
                sep = s.get(f"{{{_M_NS}}}val", "")
        return beg, end, sep

    for tag in [f"{{{_M_NS}}}sup", f"{{{_M_NS}}}sub"]:
        for script_elem in root.iter(tag):
            to_replace = [
                (i, c) for i, c in enumerate(script_elem)
                if c.tag == f"{{{_M_NS}}}d"
            ]
            for i, d_elem in reversed(to_replace):
                beg, end, sep = get_delimiter_info(d_elem)
                # 2026-05-10: 多组 <m:e> 之间需插入 sepChr（如 (t-k) 中的 −），否则分隔符丢失
                e_groups = [list(e_child) for e_child in d_elem.findall(f"{{{_M_NS}}}e")]
                content = []
                for gi, group in enumerate(e_groups):
                    content.extend(group)
                    if gi < len(e_groups) - 1 and sep:
                        content.append(make_run(sep))
                script_elem.remove(d_elem)
                new_elems = ([make_run(beg)] if beg else []) + content + ([make_run(end)] if end else [])
                for j, elem in enumerate(new_elems):
                    script_elem.insert(i + j, elem)


def latex_to_omath(latex: str) -> str:
    mathml_str = latex2mathml.converter.convert(latex)
    mathml_doc = etree.fromstring(mathml_str.encode("utf-8"))
    transform = _get_transform()
    omml_doc = transform(mathml_doc)
    _fix_script_delimiters(omml_doc.getroot())
    return etree.tostring(omml_doc, encoding="unicode")


def scan_and_convert_dollar_inline(elements: list) -> list:
    """扫描所有元素中的 $...$ 行内公式，直接转换，不经过 LLM。失败则跳过该公式。"""
    results = []
    for elem in elements:
        content = elem.get('content', '')
        matches = list(_DOLLAR_INLINE.finditer(content))
        if not matches:
            continue
        segments = []
        last_end = 0
        for match in matches:
            latex = match.group(1).strip()
            omath = ''
            try:
                omath = latex_to_omath(latex)
            except Exception:
                pass
            segments.append({
                'text_before': content[last_end:match.start()],
                'omath': omath,
                'text_after': '',
            })
            last_end = match.end()
        segments[-1]['text_after'] = content[last_end:]
        elem_id = elem.get('id')
        for seg in segments:
            results.append({
                'id': elem_id,
                'text_before': seg['text_before'],
                'omath': seg['omath'],
                'text_after': seg['text_after'],
                'label': '',
            })
    return results


def convert_formula_list(formula_items: list) -> list:
    results = []
    for item in formula_items:
        latex = item.get("latex_formula", "")
        omath = ""
        error = None
        if latex:
            try:
                omath = latex_to_omath(latex)
            except Exception as e:
                error = str(e)
        results.append({
            "id": item.get("id"),
            "text_before": item.get("text_before", ""),
            "omath": omath,
            "text_after": item.get("text_after", ""),
            "label": item.get("label", ""),
            **({"error": error} if error else {}),
        })
    return results
