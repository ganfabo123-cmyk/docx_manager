import glob
from dataclasses import dataclass
from pathlib import Path
from typing import Union

_XSLT_TRANSFORM = None


@dataclass
class InlineFormula:
    """段落内行内公式，伴随前后文本"""
    text_before: str
    formula: str      # LaTeX 格式
    text_after: str


@dataclass
class BlockFormula:
    """独立公式，带编号标注（如 "4-2"，位于公式结尾），不属于段落行内"""
    label: str
    formula: str      # LaTeX 格式


FormulaItem = Union[InlineFormula, BlockFormula]


def _find_mml2omml_xsl() -> str:
    local = Path(__file__).parent / "MML2OMML.XSL"
    if local.exists():
        return str(local)
    patterns = [
        r"C:\Program Files\Microsoft Office\root\Office*\MML2OMML.XSL",
        r"C:\Program Files (x86)\Microsoft Office\root\Office*\MML2OMML.XSL",
        r"C:\Program Files\Microsoft Office\Office*\MML2OMML.XSL",
    ]
    for pattern in patterns:
        matches = glob.glob(pattern)
        if matches:
            return matches[0]
    raise FileNotFoundError(
        "未找到 MML2OMML.XSL。请将其放置在 formula_generate.py 同级目录，\n"
        r"或确认 Microsoft Office 已安装（通常位于 C:\Program Files\Microsoft Office\root\OfficeXX\MML2OMML.XSL）"
    )


def _get_transform():
    global _XSLT_TRANSFORM
    if _XSLT_TRANSFORM is None:
        from lxml import etree
        _XSLT_TRANSFORM = etree.XSLT(etree.parse(_find_mml2omml_xsl()))
    return _XSLT_TRANSFORM


def _latex_to_omml(latex_str: str) -> str:
    import latex2mathml.converter
    from lxml import etree
    mml = latex2mathml.converter.convert(latex_str)
    mml_tree = etree.fromstring(mml.encode("utf-8"))
    omml_tree = _get_transform()(mml_tree)
    return etree.tostring(omml_tree, encoding="unicode")


def convert(formula_items: list[FormulaItem]) -> list[dict]:
    result = []
    for item in formula_items:
        omml = _latex_to_omml(item.formula)
        if isinstance(item, InlineFormula):
            block = {
                "type":      "formula",
                "label":     "",
                "omml":      omml,
                "is_inline": True,
            }
            if item.text_before:
                block["text_before"] = item.text_before
            if item.text_after:
                block["text_after"] = item.text_after
        else:
            block = {
                "type":  "formula",
                "label": item.label,
                "omml":  omml,
            }
        result.append(block)
    return result


def generate(raw_text: str) -> list[dict]:
    from llm_router import route_formula
    return convert(route_formula(raw_text))
