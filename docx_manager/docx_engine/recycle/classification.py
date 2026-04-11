"""
classification.py
=================
Analyses data/extraction.json and assigns a semantic class to every body_element.

The goal (per REVIEW.md) is to identify the distinct *types* of elements so that
semantic.json can be designed around those types.

Semantic classes
----------------
Surface type  → semantic class(es)
──────────────────────────────────────────────────────────────────────────────
paragraph     → HEADING_1 / HEADING_2 / HEADING_3   (style 2 / 3 / 4)
                COVER_TEXT                            (style 12, no section_break)
                ABSTRACT_BODY                         (style 24, no section_break)
                REFERENCE_ENTRY                       (style 9)
                ACKNOWLEDGEMENT                       (style 7)
                IMAGE_PARA                            (any run has drawing_xml)
                FORMULA_OLE                           (any run has object_xml)
                SECTION_BOUNDARY                      (has section_break)
                BODY_TEXT                             (style None/other, has text)
                EMPTY_SPACER                          (no text, no drawing, no object)
raw_xml       → TOC_BLOCK                            (instrText starts with TOC)
                TOC_ENTRY                             (instrText starts with HYPERLINK,
                                                       inside same TOC field block)
                FIELD_BLOCK                           (any other preserved field)
table         → TABLE
bookmarkEnd   → BOOKMARK_END
──────────────────────────────────────────────────────────────────────────────

Output
------
Prints a summary report to stdout.
Writes data/classification.json with each element annotated:
    { ...original element..., "_class": "<CLASS>" }
"""

import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

# ── Paths ────────────────────────────────────────────────────────────────────

INPUT_PATH  = Path("data/extraction.json")
OUTPUT_PATH = Path("data/classification.json")

# ── Style-ID → heading level map (HIT template) ─────────────────────────────

HEADING_STYLE_LEVELS = {
    "2": 1,   # 第1章  Introduction
    "3": 2,   # 1.1  Sub-section
    "4": 3,   # 1.1.1  Sub-sub-section
}

COVER_STYLES    = {"12"}
ABSTRACT_STYLES = {"24"}
REFERENCE_STYLES = {"9"}
ACK_STYLES      = {"7"}

# ── Helpers ──────────────────────────────────────────────────────────────────

def _has_drawing(para: dict) -> bool:
    return any(r.get("drawing_xml") for r in para.get("runs", []))

def _has_object(para: dict) -> bool:
    return any(r.get("object_xml") for r in para.get("runs", []))

def _raw_xml_instr(elem: dict) -> str:
    """Return the first instrText word (upper-cased) from a raw_xml element."""
    xml = elem.get("xml", "")
    parts = re.findall(r"instrText[^>]*>([^<]*)", xml)
    joined = " ".join(parts).strip()
    words = joined.split()
    return words[0].upper() if words else ""

def _classify_raw(elem: dict) -> str:
    first = _raw_xml_instr(elem)
    if first == "TOC":
        return "TOC_BLOCK"
    if first == "HYPERLINK":
        return "TOC_ENTRY"
    if first:
        return "FIELD_BLOCK"
    # Empty instrText — likely a TOC filler/continuation paragraph
    return "TOC_ENTRY"

def _classify_paragraph(para: dict) -> str:
    style = para.get("style")
    sb    = para.get("section_break")
    text  = (para.get("text") or "").strip()

    # Section boundary wins — these paragraphs contain sectPr markup
    if sb:
        return "SECTION_BOUNDARY"

    # Style-based heading detection
    if style in HEADING_STYLE_LEVELS:
        level = HEADING_STYLE_LEVELS[style]
        return f"HEADING_{level}"

    # Cover page
    if style in COVER_STYLES:
        return "COVER_TEXT"

    # Abstract / body text (style 24)
    if style in ABSTRACT_STYLES:
        return "ABSTRACT_BODY"

    # Reference entries
    if style in REFERENCE_STYLES:
        return "REFERENCE_ENTRY"

    # Acknowledgement
    if style in ACK_STYLES:
        return "ACKNOWLEDGEMENT"

    # Content-based for unstyled paragraphs
    if _has_drawing(para):
        return "IMAGE_PARA"
    if _has_object(para):
        return "FORMULA_OLE"

    if text:
        return "BODY_TEXT"

    return "EMPTY_SPACER"

def classify(elem: dict) -> str:
    t = elem.get("type")
    if t == "paragraph":
        return _classify_paragraph(elem)
    if t == "raw_xml":
        return _classify_raw(elem)
    if t == "table":
        return "TABLE"
    if t == "bookmarkEnd":
        return "BOOKMARK_END"
    return "UNKNOWN"

# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    if not INPUT_PATH.exists():
        print(f"ERROR: {INPUT_PATH} not found", file=sys.stderr)
        sys.exit(1)

    with open(INPUT_PATH, encoding="utf-8") as f:
        data = json.load(f)

    body = data.get("body_elements", [])

    # Annotate
    annotated = []
    class_counter: Counter = Counter()
    by_class: defaultdict = defaultdict(list)
    first_per_class: dict = {}  # 保存每类第一个元素

    for elem in body:
        cls = classify(elem)
        annotated_elem = dict(elem)
        annotated_elem["_class"] = cls
        annotated.append(annotated_elem)
        class_counter[cls] += 1
        by_class[cls].append(elem)
        # 只记录每类第一个元素
        if cls not in first_per_class:
            first_per_class[cls] = elem

    # ── Print report ────────────────────────────────────────────────────────

    total = len(body)
    print("=" * 60)
    print(f"  EXTRACTION.JSON  —  Semantic Classification Report")
    print("=" * 60)
    print(f"  Total body elements : {total}")
    print(f"  Distinct classes    : {len(class_counter)}")
    print()

    # Class summary table
    print(f"  {'CLASS':<20} {'COUNT':>6}  {'%':>6}  DESCRIPTION")
    print(f"  {'-'*20}  {'-'*6}  {'-'*6}  {'-'*30}")

    descriptions = {
        "HEADING_1"       : "Chapter headings (style 2)",
        "HEADING_2"       : "Section headings (style 3)",
        "HEADING_3"       : "Sub-section headings (style 4)",
        "COVER_TEXT"      : "Cover page text (style 12)",
        "ABSTRACT_BODY"   : "Abstract / body text (style 24)",
        "REFERENCE_ENTRY" : "Reference list entries (style 9)",
        "ACKNOWLEDGEMENT" : "Acknowledgement text (style 7)",
        "IMAGE_PARA"      : "Paragraph containing drawing/image run",
        "FORMULA_OLE"     : "Paragraph containing OLE formula object",
        "SECTION_BOUNDARY": "Paragraph with w:sectPr (section break)",
        "BODY_TEXT"       : "Unstyled paragraph with text content",
        "EMPTY_SPACER"    : "Blank paragraph (spacing / placeholder)",
        "TOC_BLOCK"       : "raw_xml: TOC field (fldChar begin→end)",
        "TOC_ENTRY"       : "raw_xml: HYPERLINK entry inside TOC block",
        "FIELD_BLOCK"     : "raw_xml: other preserved field structure",
        "TABLE"           : "Table element",
        "BOOKMARK_END"    : "Orphan bookmarkEnd marker",
        "UNKNOWN"         : "Unrecognised element",
    }

    for cls, count in sorted(class_counter.items(), key=lambda x: -x[1]):
        pct = 100.0 * count / total
        desc = descriptions.get(cls, "")
        print(f"  {cls:<20} {count:>6}  {pct:>5.1f}%  {desc}")

    print()

    # ── Per-class examples (only one per class) ────────────────────────────

    print("=" * 60)
    print("  EXAMPLES PER CLASS (1 ELEMENT EACH)")
    print("=" * 60)

    for cls in sorted(first_per_class.keys()):
        elem = first_per_class[cls]
        idx   = elem.get("index", "?")
        etype = elem.get("type")
        print(f"\n── {cls} ──")
        if etype == "paragraph":
            style = elem.get("style")
            text  = (elem.get("text") or "").strip()[:70]
            sb    = bool(elem.get("section_break"))
            dr    = _has_drawing(elem)
            obj   = _has_object(elem)
            flags = []
            if sb:  flags.append("section_break")
            if dr:  flags.append("drawing")
            if obj: flags.append("object")
            flag_str = f"  [{', '.join(flags)}]" if flags else ""
            print(f"  [{idx:>3}] para  style={style!r:4}  {text!r}{flag_str}")
        elif etype == "raw_xml":
            instr = _raw_xml_instr(elem)
            print(f"  [{idx:>3}] raw_xml  instrText_first={instr!r}")
        elif etype == "table":
            nrows = len(elem.get("rows", []))
            ncols = len((elem.get("rows") or [[]])[0])
            print(f"  [{idx:>3}] table  {nrows}×{ncols}")
        else:
            print(f"  [{idx:>3}] {etype}")

    # ── Save output ─────────────────────────────────────────────────────────

    output_data = dict(data)
    # 只写每类第一个元素
    output_data["body_elements"] = list(first_per_class.values())
    output_data["_classification_summary"] = {
        "total": total,
        "classes": {cls: count for cls, count in class_counter.most_common()},
    }

    OUTPUT_PATH.parent.mkdir(exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f"  Annotated JSON (1 element per class) written to: {OUTPUT_PATH}")
    print()

if __name__ == "__main__":
    main()
