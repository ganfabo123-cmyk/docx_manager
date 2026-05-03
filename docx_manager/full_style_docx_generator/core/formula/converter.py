import re
from pathlib import Path
import latex2mathml.converter
from lxml import etree

_DOLLAR_INLINE = re.compile(r'\$([^$\n]+?)\$')

_XSL_PATH = Path(__file__).parent / "assets" / "MML2OMML.XSL"
_transform = None


def _get_transform():
    global _transform
    if _transform is None:
        xslt_doc = etree.parse(str(_XSL_PATH))
        _transform = etree.XSLT(xslt_doc)
    return _transform


def latex_to_omath(latex: str) -> str:
    mathml_str = latex2mathml.converter.convert(latex)
    mathml_doc = etree.fromstring(mathml_str.encode("utf-8"))
    transform = _get_transform()
    omml_doc = transform(mathml_doc)
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
