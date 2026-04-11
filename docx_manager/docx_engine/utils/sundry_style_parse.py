"""
sundry_style_parse.py — 杂项样式解析器

解析 data/extraction.json 中以下三类元素的格式与内容:
  1. 中文关键词段落  (关键词：…)
  2. 英文关键词段落  (Keywords：…)
  3. 正文引用标注    (行内上标 [n] 的段落)
  4. 表格            (表头 + 数据行，含中英文混排与上下标)

输出: data/sundry_parsed.json
"""

import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
DATA_DIR   = BASE_DIR / "data"
INPUT_JSON = DATA_DIR / "extraction.json"
OUTPUT_JSON = DATA_DIR / "sundry_parsed.json"

# ── OOXML namespace ────────────────────────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ── XML helpers ───────────────────────────────────────────────────────────────

def _parse_xml(xml_str: Optional[str]) -> Optional[ET.Element]:
    if not xml_str:
        return None
    try:
        return ET.fromstring(xml_str)
    except ET.ParseError:
        return None


def _w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _get(el: ET.Element, tag: str) -> Optional[str]:
    """Return w:val of a direct child tag, or None."""
    child = el.find(_w(tag))
    if child is None:
        return None
    return child.get(_w("val"))


# ── Run property extraction ────────────────────────────────────────────────────

def _parse_rpr(rpr_str: Optional[str]) -> dict:
    """
    Parse a run's rPr XML string into a flat property dict.

    Returned keys (all optional / None when absent):
        bold        bool
        italic      bool
        underline   bool
        vert_align  "superscript" | "subscript" | None
        font_cn     str   (eastAsia / ascii when Chinese-looking)
        font_en     str   (ascii / hAnsi when Latin-looking)
        size_pt     float (w:sz val / 2)
        kern        int   (w:kern val)
    """
    props: dict = {
        "bold": False,
        "italic": False,
        "underline": False,
        "vert_align": None,
        "font_cn": None,
        "font_en": None,
        "size_pt": None,
        "kern": None,
    }
    root = _parse_xml(rpr_str)
    if root is None:
        return props

    # Bold / italic / underline — presence of tag means true
    props["bold"]      = root.find(_w("b"))   is not None
    props["italic"]    = root.find(_w("i"))   is not None or root.find(_w("iCs")) is not None
    props["underline"] = root.find(_w("u"))   is not None

    # Vertical alignment
    va = root.find(_w("vertAlign"))
    if va is not None:
        props["vert_align"] = va.get(_w("val"))

    # Fonts
    fonts_el = root.find(_w("rFonts"))
    if fonts_el is not None:
        ea = fonts_el.get(_w("eastAsia"))
        ascii_ = fonts_el.get(_w("ascii"))
        hAnsi  = fonts_el.get(_w("hAnsi"))
        props["font_cn"] = ea or ascii_
        props["font_en"] = ascii_ or hAnsi

    # Size in pt (w:sz val is in half-points)
    sz = root.find(_w("sz"))
    if sz is not None:
        try:
            props["size_pt"] = int(sz.get(_w("val"), 0)) / 2
        except (ValueError, TypeError):
            pass

    # Kern
    kern = root.find(_w("kern"))
    if kern is not None:
        try:
            props["kern"] = int(kern.get(_w("val"), 0))
        except (ValueError, TypeError):
            pass

    return props


def _annotate_run(run: dict) -> dict:
    """Return run dict augmented with parsed rPr properties."""
    rpr = _parse_rpr(run.get("rPr"))
    return {
        "text": run.get("text", ""),
        **rpr,
    }


# ── Keyword parsing ────────────────────────────────────────────────────────────

def _is_chinese_keywords(para: dict) -> bool:
    text = (para.get("text") or "").lstrip()
    return text.startswith("关键词")


def _is_english_keywords(para: dict) -> bool:
    text = (para.get("text") or "").lstrip()
    return text.startswith("Keywords")


def _extract_keyword_items(text: str, prefix: str, sep: str) -> list[str]:
    """
    Strip *prefix* (and an optional following punctuation char) from *text*,
    then split the remainder by *sep*, stripping whitespace from each item.
    """
    body = text.lstrip()
    # Remove prefix and leading punctuation (：: etc.)
    body = re.sub(r"^" + re.escape(prefix) + r"\s*[：:]\s*", "", body)
    items = [kw.strip() for kw in body.split(sep) if kw.strip()]
    return items


def parse_keywords(body_elements: list[dict]) -> dict:
    """
    Locate the Chinese and English keyword paragraphs and extract them.

    Returns:
        {
          "chinese": { "index": int, "raw_text": str, "keywords": [...],
                       "runs": [...annotated runs...] },
          "english": { ... }
        }
    """
    result: dict = {"chinese": None, "english": None}

    for elem in body_elements:
        if elem.get("type") != "paragraph":
            continue

        if result["chinese"] is None and _is_chinese_keywords(elem):
            raw = elem.get("text", "")
            result["chinese"] = {
                "index":    elem.get("index"),
                "raw_text": raw,
                "keywords": _extract_keyword_items(raw, "关键词", "；"),
                "runs":     [_annotate_run(r) for r in elem.get("runs", [])],
            }

        elif result["english"] is None and _is_english_keywords(elem):
            raw = elem.get("text", "")
            result["english"] = {
                "index":    elem.get("index"),
                "raw_text": raw,
                "keywords": _extract_keyword_items(raw, "Keywords", ","),
                "runs":     [_annotate_run(r) for r in elem.get("runs", [])],
            }

        if result["chinese"] and result["english"]:
            break

    return result


# ── Citation parsing ───────────────────────────────────────────────────────────

_CITATION_RE = re.compile(r"\[(\d+(?:[,，]\d+)*)\]")


def _extract_citations_from_run(run_text: str) -> list[str]:
    """Return list of citation tokens (e.g. '[3]', '[1,2]') found in text."""
    return _CITATION_RE.findall(run_text)


def _parse_citation_paragraph(para: dict) -> Optional[dict]:
    """
    If *para* contains any in-line superscript citation runs, return a
    structured record; otherwise return None.

    A citation run is one whose rPr contains  <w:vertAlign w:val="superscript"/>
    AND whose text matches the [n] pattern.
    """
    citations: list[dict] = []
    annotated_runs: list[dict] = []

    for run in para.get("runs", []):
        arpr = _parse_rpr(run.get("rPr"))
        text = run.get("text", "")
        ann  = {"text": text, **arpr}
        annotated_runs.append(ann)

        if arpr.get("vert_align") == "superscript":
            refs = _extract_citations_from_run(text)
            if refs:
                for ref in refs:
                    citations.append({"ref": f"[{ref}]", "superscript_text": text})

    if not citations:
        return None

    return {
        "index":    para.get("index"),
        "style":    para.get("style"),
        "text":     para.get("text", ""),
        "citations": citations,
        "runs":     annotated_runs,
    }


def parse_citations(body_elements: list[dict]) -> list[dict]:
    """
    Return all body paragraphs that contain at least one in-text superscript
    citation (format [n]).
    """
    results = []
    for elem in body_elements:
        if elem.get("type") != "paragraph":
            continue
        record = _parse_citation_paragraph(elem)
        if record:
            results.append(record)
    return results


# ── Table parsing ─────────────────────────────────────────────────────────────

def _annotate_cell_paragraph(para: dict) -> dict:
    """Return a simplified representation of a cell paragraph."""
    return {
        "index":   para.get("index"),
        "style":   para.get("style"),
        "text":    para.get("text", ""),
        "runs":    [_annotate_run(r) for r in para.get("runs", [])],
    }


def _annotate_cell(cell: dict) -> dict:
    return {
        "text":       cell.get("text", ""),
        "paragraphs": [_annotate_cell_paragraph(p) for p in cell.get("paragraphs", [])],
    }


def parse_tables(body_elements: list[dict]) -> list[dict]:
    """
    Parse all table elements and return their annotated structure.

    Each table entry:
        {
          "index": int,
          "rows": [
            [ { "text": str, "paragraphs": [...] }, ... ],  # one cell per entry
            ...
          ]
        }
    """
    results = []
    for elem in body_elements:
        if elem.get("type") != "table":
            continue
        annotated_rows = []
        for row in elem.get("rows", []):
            annotated_rows.append([_annotate_cell(cell) for cell in row])
        results.append({
            "index": elem.get("index"),
            "rows":  annotated_rows,
        })
    return results


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    if not INPUT_JSON.exists():
        raise FileNotFoundError(f"extraction.json not found at {INPUT_JSON}")

    with open(INPUT_JSON, encoding="utf-8") as fh:
        data = json.load(fh)

    body: list[dict] = data.get("body_elements", [])

    output = {
        "keywords":  parse_keywords(body),
        "citations": parse_citations(body),
        "tables":    parse_tables(body),
    }

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as fh:
        json.dump(output, fh, ensure_ascii=False, indent=2)

    # ── Summary ───────────────────────────────────────────────────────────────
    kw = output["keywords"]
    cn = kw["chinese"]
    en = kw["english"]

    print("=== Sundry Style Parse ===")
    if cn:
        print(f"Chinese keywords  (para {cn['index']}): {cn['keywords']}")
    else:
        print("Chinese keywords: NOT FOUND")

    if en:
        print(f"English keywords  (para {en['index']}): {en['keywords']}")
    else:
        print("English keywords: NOT FOUND")

    print(f"Citation paragraphs : {len(output['citations'])}")
    print(f"Tables              : {len(output['tables'])}")
    print(f"\nOutput written to {OUTPUT_JSON}")


if __name__ == "__main__":
    main()
