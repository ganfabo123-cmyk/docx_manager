from pathlib import Path
import latex2mathml.converter
from lxml import etree

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
