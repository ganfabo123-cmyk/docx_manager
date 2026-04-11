"""
formatting_style.py — Run-level Style Normalizer for DOCX Engine V3

What it does
────────────
1. Copies template/ → template_normalized/ (original never touched).
2. Leaves every paragraph's `pPr` and `style` field EXACTLY as-is.
3. For every run (body_elements, headers, footers):
   • Parses its rPr XML.
   • Finds or creates a character style in styles.xml that captures those
     run properties exactly.
   • Rewrites the run's rPr to just  <w:rPr><w:rStyle w:val="…"/></w:rPr>
     (or null when the run carries no formatting of its own).
4. Appends new character styles to template_normalized/word/styles.xml.
5. Writes data/extraction_normalized.json.
6. Writes data/style_lineage.json  — a human-readable record of every new
   character style: its name, properties, parent paragraph contexts, and
   the XML of those parent paragraph styles.

Style naming convention
───────────────────────
New character styles are named from their dominant rPr properties:
  primary font  +  size in pt  +  Bold / Italic / Underline flags
  e.g.  "隶书 26pt", "楷体_GB2312 16pt Bold", "Times New Roman 12pt Italic"
A numeric suffix (_1, _2, …) is appended only when two distinct rPr
combinations produce the same descriptive base name.
"""

import copy
import json
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Optional
import shutil

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR     = Path(__file__).parent
DATA_DIR     = BASE_DIR / "data"
TEMPLATE_DIR = BASE_DIR / "template"
NORM_DIR     = BASE_DIR / "template_normalized"
INPUT_JSON   = DATA_DIR / "extraction.json"
OUTPUT_JSON  = DATA_DIR / "extraction_normalized.json"
LINEAGE_JSON = DATA_DIR / "style_lineage.json"
STYLES_PATH  = NORM_DIR / "word" / "styles.xml"

# ── OOXML namespace ────────────────────────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_NS_MAP: dict[str, str] = {
    "w":             W,
    "wpc":           "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "mc":            "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o":             "urn:schemas-microsoft-com:office:office",
    "r":             "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m":             "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v":             "urn:schemas-microsoft-com:vml",
    "wp14":          "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp":            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w14":           "http://schemas.microsoft.com/office/word/2010/wordml",
    "w10":           "urn:schemas-microsoft-com:office:word",
    "w15":           "http://schemas.microsoft.com/office/word/2012/wordml",
    "sl":            "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "wpsCustomData": "http://www.wps.cn/officeDocument/2013/wpsCustomData",
    "a":             "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic":           "http://schemas.openxmlformats.org/drawingml/2006/picture",
}
for _p, _u in _NS_MAP.items():
    ET.register_namespace(_p, _u)

# ── Chinese font → ASCII abbreviation map (for style IDs) ─────────────────────
_FONT_ABBREV: dict[str, str] = {
    "隶书":          "Lishu",
    "方正隶书_GBK":  "FZLishu",
    "楷体":          "Kaiti",
    "楷体_GB2312":   "Kaiti",
    "黑体":          "Heiti",
    "仿宋":          "Fangsong",
    "仿宋_GB2312":   "Fangsong",
    "宋体":          "Songti",
    "微软雅黑":      "MicrosoftYaHei",
    "华文宋体":      "HWSongti",
    "华文楷体":      "HWKaiti",
}

# ── Named colours (upper-case hex → English name) ─────────────────────────────
_COLOR_NAMES: dict[str, str] = {
    "000000": "Black",
    "FFFFFF": "White",
    "FF0000": "Red",
    "0000FF": "Blue",
    "00FF00": "Green",
    "FFFF00": "Yellow",
    "FF6600": "Orange",
    "800080": "Purple",
}


# ── Low-level XML helpers ──────────────────────────────────────────────────────

def _local(tag: str) -> str:
    """{namespace}local  →  local"""
    return tag.split("}", 1)[1] if "}" in tag else tag


def _canon(el: ET.Element) -> tuple:
    """
    Canonical, hashable representation of an XML sub-tree.
    Ignores namespace URIs (local names only), sorts attributes.
    Used as a dict key for deduplication.
    """
    tag   = _local(el.tag)
    attrs = tuple(sorted((_local(k), v) for k, v in el.attrib.items()))
    text  = (el.text or "").strip()
    kids  = tuple(_canon(c) for c in el)
    return (tag, attrs, text, kids)


def _parse(xml_str: Optional[str]) -> Optional[ET.Element]:
    if not xml_str:
        return None
    try:
        return ET.fromstring(xml_str)
    except ET.ParseError:
        return None


def _ser(el: ET.Element) -> str:
    return ET.tostring(el, encoding="unicode")


# ── rPr decomposition ─────────────────────────────────────────────────────────

def _rpr_signature(rpr_el: Optional[ET.Element]) -> tuple:
    """
    Canonical tuple of a rPr's children, excluding any existing rStyle.
    Empty tuple  →  run carries no formatting of its own.
    """
    if rpr_el is None:
        return ()
    kids = [_canon(c) for c in rpr_el if _local(c.tag) != "rStyle"]
    return tuple(sorted(kids))


def _rpr_to_props(rpr_el: ET.Element) -> dict:
    """
    Convert a rPr element into a plain-dict representation for JSON output.
    Each child tag maps to a dict of its attributes.
    """
    props: dict[str, dict] = {}
    for child in rpr_el:
        loc = _local(child.tag)
        if loc == "rStyle":
            continue
        attrs = {_local(k): v for k, v in child.attrib.items()}
        props[loc] = attrs
    return props


# ── Descriptive naming ────────────────────────────────────────────────────────

def _lang_token(val: str) -> str:
    """'zh-CN' → 'zhCN',  'en-US' → 'enUS'"""
    return val.replace("-", "")


def _derive_char_display_name(rpr_el: ET.Element) -> str:
    """
    Build a human-readable display name from rPr properties.

    Priority order for name parts:
      [FontFamily | EA-Hint]  [size]  [Bold] [Italic] [Underline]
      [Superscript/Subscript] [Color] [Spc{n}] [Kern{n}] [Pos{n}] [LangTokens]

    "EA" is added when rFonts only has a hint="eastAsia" attribute (no family).
    "CS" is appended after the size when szCs is present (signals that the
    complex-script size is also set), allowing size-only variants to be
    distinguished without falling back to counter suffixes.

    Examples
    ────────
    {rFonts:hint=ea, sz:21, szCs:21}  →  "EA 10pt CS"
    {sz:21, szCs:21}                  →  "10pt CS"
    {rFonts:hint=ea, sz:21}           →  "EA 10pt"
    {sz:21}                           →  "10pt"
    {rFonts:hint=ea}                  →  "EA"
    {vertAlign:superscript}           →  "Superscript"
    {color:FF0000}                    →  "Red"
    {color:000000, position:-14}      →  "Black Pos-14"
    {rFonts:hint=ea, spacing:-4}      →  "EA Spc-4"
    {lang:zh-CN}                      →  "zhCN"
    """
    font_ea:    Optional[str] = None
    font_ascii: Optional[str] = None
    font_hint:  Optional[str] = None
    sz_val:     Optional[int] = None
    szcs_val:   Optional[int] = None
    is_bold      = False
    is_italic    = False
    is_underline = False
    vert_align:  Optional[str] = None
    color_val:   Optional[str] = None
    spacing_val: Optional[str] = None
    kern_val:    Optional[str] = None
    position_val: Optional[str] = None
    lang_tokens: list[str] = []

    for child in rpr_el:
        loc   = _local(child.tag)
        attrs = {_local(k): v for k, v in child.attrib.items()}

        if loc == "rFonts":
            font_ea    = attrs.get("eastAsia")
            font_ascii = attrs.get("ascii") or attrs.get("hAnsi")
            font_hint  = attrs.get("hint")

        elif loc == "sz":
            v = attrs.get("val", "")
            if v.isdigit():
                sz_val = int(v)

        elif loc == "szCs":
            v = attrs.get("val", "")
            if v.isdigit():
                szcs_val = int(v)

        elif loc == "b":
            if attrs.get("val", "1") not in ("0", "false"):
                is_bold = True

        elif loc == "i":
            if attrs.get("val", "1") not in ("0", "false"):
                is_italic = True

        elif loc == "u":
            if attrs.get("val", "none") != "none":
                is_underline = True

        elif loc == "vertAlign":
            vert_align = attrs.get("val")   # "superscript" / "subscript"

        elif loc == "color":
            v = attrs.get("val", "")
            if v and v not in ("auto", "none"):
                color_val = v.upper()

        elif loc == "spacing":
            v = attrs.get("val")
            if v:
                spacing_val = v

        elif loc == "kern":
            v = attrs.get("val")
            if v:
                kern_val = v

        elif loc == "position":
            v = attrs.get("val")
            if v:
                position_val = v

        elif loc == "lang":
            for attr_key in ("val", "eastAsia", "bidi"):
                lv = attrs.get(attr_key)
                if lv:
                    tok = _lang_token(lv)
                    if tok not in lang_tokens:
                        lang_tokens.append(tok)

    # ── Assemble parts ────────────────────────────────────────────────────────
    parts: list[str] = []

    # 1. Font family (or EA hint marker when no family name is available)
    primary_font = font_ea or font_ascii
    if primary_font:
        parts.append(primary_font)
    elif font_hint == "eastAsia":
        parts.append("EA")
    elif font_hint:
        parts.append(font_hint.upper())

    # 2. Size in whole points (sz is in half-points)
    if sz_val is not None:
        sz_pt = f"{sz_val // 2}pt"
        # Append CS marker when szCs is also set (even if equal to sz),
        # so that {sz+szCs} and {sz only} get distinct base names.
        if szcs_val is not None:
            parts.append(f"{sz_pt} CS")
        else:
            parts.append(sz_pt)
    elif szcs_val is not None:
        # szCs present but sz absent — show as complex-script size only
        parts.append(f"csz{szcs_val // 2}pt")

    # 3. Weight / style
    if is_bold:
        parts.append("Bold")
    if is_italic:
        parts.append("Italic")
    if is_underline:
        parts.append("Underline")

    # 4. Vertical alignment
    if vert_align:
        parts.append(vert_align.capitalize())   # "Superscript" / "Subscript"

    # 5. Colour
    if color_val:
        named = _COLOR_NAMES.get(color_val)
        parts.append(named if named else f"#{color_val}")

    # 6. Character spacing, kerning, vertical position
    if spacing_val:
        parts.append(f"Spc{spacing_val}")
    if kern_val:
        parts.append(f"Kern{kern_val}")
    if position_val:
        parts.append(f"Pos{position_val}")

    # 7. Language tokens (secondary; appended last)
    if lang_tokens:
        parts.append("-".join(lang_tokens))

    return " ".join(parts) if parts else "CharStyle"


def _display_name_to_id(name: str) -> str:
    """
    Convert a display name (may contain Chinese) to an ASCII-safe style ID.
    Chinese font names are replaced by their pinyin abbreviations.
    Result format: e.g. "cs-Lishu-26pt-Bold"
    """
    result = name
    for cn, en in _FONT_ABBREV.items():
        result = result.replace(cn, en)

    # Keep only word-chars, hyphens, underscores, spaces
    result = re.sub(r"[^\w\s\-]", "", result)
    # Collapse whitespace → hyphens
    result = re.sub(r"\s+", "-", result.strip())
    result = re.sub(r"-{2,}", "-", result)
    # Drop any non-ASCII that slipped through
    result = "".join(c for c in result if ord(c) < 128)

    base_id = f"cs-{result}" if result else "cs-CharStyle"
    return base_id


# ── StyleRegistry ──────────────────────────────────────────────────────────────

class StyleRegistry:
    """
    Wraps styles.xml in the normalized template folder.

    For each unique rPr signature it either reuses an existing character style
    (exact canonical match) or creates a new one with a descriptive name.

    Call `save()` to flush new styles to disk.
    `lineage_entries()` returns the data needed to build style_lineage.json.
    """

    def __init__(self, styles_path: Path) -> None:
        self._path = styles_path

        for _, (prefix, uri) in ET.iterparse(str(styles_path), events=["start-ns"]):
            ET.register_namespace(prefix, uri)

        self._tree = ET.parse(str(styles_path))
        self._root = self._tree.getroot()

        self._existing_ids:   set[str]        = set()
        self._used_names:     set[str]        = set()
        # canonical rPr tuple → styleId
        self._char_map:       dict[tuple, str] = {}
        # styleId → new style element (for newly created styles only)
        self._new_style_els:  dict[str, ET.Element] = {}
        # styleId → display name
        self._char_names:     dict[str, str]  = {}
        # styleId → rPr properties dict (for lineage JSON)
        self._char_props:     dict[str, dict] = {}
        # styleId → {para_style_id: count}
        self._char_contexts:  dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))

        self._index_existing()

    # ── indexing ───────────────────────────────────────────────────────────────

    def _index_existing(self) -> None:
        for style_el in self._root.findall(f"{{{W}}}style"):
            sid   = style_el.get(f"{{{W}}}styleId")
            stype = style_el.get(f"{{{W}}}type")
            name_el = style_el.find(f"{{{W}}}name")
            if sid:
                self._existing_ids.add(sid)
            if name_el is not None:
                n = name_el.get(f"{{{W}}}val", "")
                if n:
                    self._used_names.add(n)

            if stype == "character" and sid:
                rpr    = style_el.find(f"{{{W}}}rPr")
                r_kids = [_canon(c) for c in rpr] if rpr is not None else []
                sig    = tuple(sorted(r_kids))
                self._char_map[sig] = sid

    # ── ID / name allocation ───────────────────────────────────────────────────

    def _alloc_char(self, rpr_el: ET.Element) -> tuple[str, str]:
        """
        Return (style_id, display_name) for a brand-new character style.
        Appends _1, _2, … to both until both are unique.
        """
        base_name = _derive_char_display_name(rpr_el)
        base_id   = _display_name_to_id(base_name)

        candidate_name = base_name
        candidate_id   = base_id
        n = 1
        while (candidate_id   in self._existing_ids or
               candidate_name in self._used_names):
            candidate_name = f"{base_name}_{n}"
            candidate_id   = f"{base_id}_{n}"
            n += 1

        self._existing_ids.add(candidate_id)
        self._used_names.add(candidate_name)
        return candidate_id, candidate_name

    # ── public resolution API ─────────────────────────────────────────────────

    def resolve_char_style(
        self,
        sig:         tuple,
        rpr_el:      Optional[ET.Element],
        para_style:  Optional[str],
    ) -> Optional[str]:
        """
        Return the character styleId for this run signature.
        Returns None when sig is empty (run has no formatting of its own).

        para_style  – the paragraph's current `style` field, used only for
                      context tracking in the lineage report.
        """
        if not sig:
            return None

        if sig in self._char_map:
            sid = self._char_map[sig]
            # Track paragraph context even for existing styles
            if para_style:
                self._char_contexts[sid][para_style] += 1
            return sid

        assert rpr_el is not None
        sid, display_name = self._alloc_char(rpr_el)

        self._char_map[sig]       = sid
        self._char_names[sid]     = display_name
        self._char_props[sid]     = _rpr_to_props(rpr_el)

        style_el = self._build_char_style(sid, display_name, rpr_el)
        self._new_style_els[sid]  = style_el

        if para_style:
            self._char_contexts[sid][para_style] += 1
        return sid

    # ── save / report ─────────────────────────────────────────────────────────

    def save(self) -> None:
        """Append all new character style elements and write styles.xml."""
        for style_el in self._new_style_els.values():
            self._root.append(style_el)
        self._tree.write(
            str(self._path),
            xml_declaration=True,
            encoding="utf-8",
        )

    def new_styles_count(self) -> int:
        return len(self._new_style_els)

    def lineage_entries(
        self,
        existing_para_styles: dict[str, dict],
    ) -> list[dict]:
        """
        Build the list of lineage dicts for every newly created character style.

        existing_para_styles  – {styleId: {"name": …, "xml": …}}
                                pre-extracted from styles.xml
        """
        entries = []
        for sid, style_el in self._new_style_els.items():
            # Paragraph contexts where this character style appears
            contexts = []
            for para_sid, count in sorted(
                self._char_contexts.get(sid, {}).items(),
                key=lambda kv: -kv[1],   # most-frequent first
            ):
                info = existing_para_styles.get(para_sid, {})
                contexts.append({
                    "para_style_id":   para_sid,
                    "para_style_name": info.get("name", para_sid),
                    "occurrence_count": count,
                    "para_style_xml":  info.get("xml"),
                })

            entries.append({
                "style_id":           sid,
                "style_name":         self._char_names.get(sid, sid),
                "style_xml":          _ser(style_el),
                "based_on":           None,
                "properties":         self._char_props.get(sid, {}),
                "paragraph_contexts": contexts,
            })

        # Sort by style_name for readability
        entries.sort(key=lambda e: e["style_name"])
        return entries

    # ── style element builder ─────────────────────────────────────────────────

    @staticmethod
    def _build_char_style(
        sid:          str,
        display_name: str,
        rpr_el:       ET.Element,
    ) -> ET.Element:
        style = ET.Element(f"{{{W}}}style")
        style.set(f"{{{W}}}type",    "character")
        style.set(f"{{{W}}}styleId", sid)

        name_el = ET.SubElement(style, f"{{{W}}}name")
        name_el.set(f"{{{W}}}val", display_name)

        run_children = [
            copy.deepcopy(c) for c in rpr_el
            if _local(c.tag) != "rStyle"
        ]
        if run_children:
            new_rpr = ET.SubElement(style, f"{{{W}}}rPr")
            for c in run_children:
                new_rpr.append(c)

        return style


# ── Run normalization ─────────────────────────────────────────────────────────

def _build_rpr_str(char_style_id: Optional[str]) -> Optional[str]:
    """
    <w:rPr><w:rStyle w:val="{char_style_id}"/></w:rPr>
    Returns None when the run has no character style (inherits everything).
    """
    if not char_style_id:
        return None
    rpr = ET.Element(f"{{{W}}}rPr")
    rs  = ET.SubElement(rpr, f"{{{W}}}rStyle")
    rs.set(f"{{{W}}}val", char_style_id)
    return _ser(rpr)


def _normalize_runs(
    para:        dict,
    registry:    StyleRegistry,
) -> dict:
    """
    Return an updated copy of the paragraph dict with all run rPr values
    replaced by character style references.  pPr and `style` are untouched.
    """
    para = dict(para)
    # Use the paragraph's style ID as context; fall back to "1" (Normal)
    # so that runs in unstyled paragraphs still appear in the lineage.
    para_style = para.get("style") or "1"

    new_runs = []
    for run in para.get("runs") or []:
        run    = dict(run)
        rpr_el = _parse(run.get("rPr"))
        sig    = _rpr_signature(rpr_el)
        char_id = registry.resolve_char_style(sig, rpr_el, para_style)
        run["rPr"] = _build_rpr_str(char_id)
        new_runs.append(run)

    para["runs"] = new_runs
    return para


def _normalize_para_list(paras: list, registry: StyleRegistry) -> list:
    return [_normalize_runs(p, registry) for p in paras]


# ── Paragraph-style index (for lineage report) ────────────────────────────────

def _index_para_styles(styles_path: Path) -> dict[str, dict]:
    """
    Return {styleId: {"name": str, "xml": str}} for every paragraph style
    found in styles.xml.  Used to enrich the lineage JSON.
    """
    result: dict[str, dict] = {}
    tree = ET.parse(str(styles_path))
    root = tree.getroot()

    for style_el in root.findall(f"{{{W}}}style"):
        if style_el.get(f"{{{W}}}type") != "paragraph":
            continue
        sid = style_el.get(f"{{{W}}}styleId")
        if not sid:
            continue
        name_el = style_el.find(f"{{{W}}}name")
        name    = name_el.get(f"{{{W}}}val", sid) if name_el is not None else sid
        result[sid] = {
            "name": name,
            "xml":  _ser(style_el),
        }
    return result


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    # ── 1. Duplicate the template folder ──────────────────────────────────────
    if NORM_DIR.exists():
        shutil.rmtree(NORM_DIR)
    shutil.copytree(TEMPLATE_DIR, NORM_DIR)
    print(f"[1/6] Copied template/  →  {NORM_DIR.name}/")

    # ── 2. Load extraction.json ───────────────────────────────────────────────
    with open(INPUT_JSON, encoding="utf-8") as fh:
        extraction = json.load(fh)

    n_body = len(extraction.get("body_elements", []))
    n_hdrs = sum(
        len(v.get("paragraphs", v) if isinstance(v, dict) else v)
        for v in extraction.get("headers", {}).values()
    )
    n_ftrs = sum(
        len(v.get("paragraphs", v) if isinstance(v, dict) else v)
        for v in extraction.get("footers", {}).values()
    )
    print(
        f"[2/6] Loaded {INPUT_JSON.name}  "
        f"({n_body} body, {n_hdrs} header, {n_ftrs} footer paragraphs)"
    )

    # ── 3. Build style registry ───────────────────────────────────────────────
    registry = StyleRegistry(STYLES_PATH)
    print(f"[3/6] Indexed {len(registry._existing_ids)} existing styles")

    # ── 4. Normalize runs (pPr / style fields are never touched) ──────────────
    extraction["body_elements"] = _normalize_para_list(
        extraction.get("body_elements", []), registry
    )

    for hdr_name, hdr_data in extraction.get("headers", {}).items():
        if isinstance(hdr_data, dict) and "paragraphs" in hdr_data:
            hdr_data["paragraphs"] = _normalize_para_list(
                hdr_data["paragraphs"], registry
            )
        elif isinstance(hdr_data, list):
            extraction["headers"][hdr_name] = _normalize_para_list(
                hdr_data, registry
            )

    for ftr_name, ftr_data in extraction.get("footers", {}).items():
        if isinstance(ftr_data, dict) and "paragraphs" in ftr_data:
            ftr_data["paragraphs"] = _normalize_para_list(
                ftr_data["paragraphs"], registry
            )
        elif isinstance(ftr_data, list):
            extraction["footers"][ftr_name] = _normalize_para_list(
                ftr_data, registry
            )

    print(f"[4/6] Created {registry.new_styles_count()} new character styles")

    # ── 5. Write outputs ──────────────────────────────────────────────────────
    registry.save()
    print(f"      Wrote styles  →  {STYLES_PATH}")

    with open(OUTPUT_JSON, "w", encoding="utf-8") as fh:
        json.dump(extraction, fh, ensure_ascii=False, indent=2)
    print(f"[5/6] Wrote {OUTPUT_JSON.name}")

    # ── 6. Write style lineage JSON ───────────────────────────────────────────
    # Re-index paragraph styles *after* save so the XML includes new styles too.
    # But we want only the original paragraph styles for context, so index
    # from the normalized file (new styles are character styles, not paragraph,
    # so this is fine either way).
    para_style_index = _index_para_styles(STYLES_PATH)

    lineage = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source":        INPUT_JSON.name,
        "total_new_character_styles": registry.new_styles_count(),
        "character_styles": registry.lineage_entries(para_style_index),
    }

    with open(LINEAGE_JSON, "w", encoding="utf-8") as fh:
        json.dump(lineage, fh, ensure_ascii=False, indent=2)
    print(f"[6/6] Wrote {LINEAGE_JSON.name}")


if __name__ == "__main__":
    main()
