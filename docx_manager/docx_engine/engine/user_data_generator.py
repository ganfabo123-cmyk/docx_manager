"""
user_data_generator.py

读取 data/full_parsed.json (纯内容数据) 和 sections_config/hit_config.json (格式配置)，
生成可供 user_data_compiler.py 编译的 data/user_data.json。

── full_parsed.json 元素格式 ─────────────────────────────────────────────────
显式格式（推荐）:
    {"type": "abstract_cn", "data": {"content": "...", "keywords": [...]}}
    {"type": "abstract_en", "data": {"content": "...", "keywords": [...]}}
    {"type": "h1",       "data": "章节标题"}
    {"type": "h2",       "data": "节标题"}
    {"type": "h3",       "data": "小节标题"}
    {"type": "h4",       "data": "小小节标题"}
    {"type": "text",     "data": "正文段落"}
    {"type": "figure",   "data": {"drawing_xml": "...", "caption": "..."}}
    {"type": "table",    "data": {"rows": [[...], ...], "caption": "..."}}
    {"type": "equation", "data": {"expression": "...", "suffix": "(1-1)"}}
    {"type": "reference","data": {"index": 1, "content": "...",
                                   "before": "...", "after": "..."}}

Markdown 兼容格式（兼容当前 full_parsed.json 占位文件）:
    {"type": "body", "content": "# 一级标题"}   → h1
    {"type": "body", "content": "## 二级标题"}  → h2
    {"type": "body", "content": "### 三级标题"} → h3
    {"type": "body", "content": "正文文本"}      → text

── hit_config.json 关键配置说明 ─────────────────────────────────────────────
  dismiss_before_front   : 若为 true，解析到第一个前言标题前的元素全部忽略
  front_matter           : 前言节列表（摘要/Abstract/目录），各含 section_break 配置
                           auto_generate_toc=true 的节自动在标题后插入 generate_toc
  body_start_section_break : 目录末尾分节配置，含 restart_page_number=1
  body_section_breaks    : 正文各节分节配置池，按一级标题顺序依次消费；
                           池耗尽后追加空分节（继承前节配置）
  final_section_break    : 文档末尾分节配置

── 生成逻辑 ─────────────────────────────────────────────────────────────────
  1. dismiss_before_front → 丢弃前言第一标题之前的所有元素
  2. 分离前言元素（摘要/Abstract 标题及其下属文本）与正文元素
  3. 用前言文本构造 insert_abstract_with_keywords 动作
  4. 依次输出各前言节的 insert_section_break；auto_generate_toc 节额外输出 generate_toc
  5. 输出 body_start_section_break
  6. 遍历正文元素：每遇到一级标题（第二个起）先输出 body_section_breaks 池中的下一条
  7. 末尾输出 final_section_break

Usage:
    python user_data_generator.py [full_parsed.json] [hit_config.json] [output.json]
"""

import json
import re
import sys
from pathlib import Path

_BASE        = Path(__file__).parent.parent
_FULL_PARSED = _BASE / "data" / "full_parsed.json"
_HIT_CONFIG  = _BASE / "sections_config" / "hit_config.json"
_OUTPUT      = _BASE / "data" / "user_data.json"

_MD_HEADING  = re.compile(r'^(#{1,4})\s+(.*)')
_KW_CN_RE    = re.compile(r'^关键词\s*[：:]\s*')
_KW_EN_RE    = re.compile(r'^[Kk]eywords\s*[：:]\s*')


# ── Element normalisation ──────────────────────────────────────────────────────

def _normalize(raw: dict) -> dict:
    """
    Convert a raw element from full_parsed.json to canonical form:
        {"type": <str>, "data": <str|dict>}

    Supports three source formats:
      1. Explicit {type, data}  — returned as-is (canonical form).
      2. Structured {type, content, style?} where type is a known non-text type
         — content mapped to data with type-specific logic.
      3. Markdown-style {type:"body", content:"# Heading" | "text…"}
         — heading prefix detected and mapped to h1-h4 / text.

    Type aliases accepted (in addition to canonical names):
        heading1/2/3/4  → h1/h2/h3/h4
        image           → figure
        formula         → equation
    """
    typ = raw.get("type", "text")

    # ── Type aliases ──────────────────────────────────────────────────────────
    _HEADING_ALIAS = {"heading1": "h1", "heading2": "h2",
                      "heading3": "h3", "heading4": "h4"}
    if typ in _HEADING_ALIAS:
        # content is always a plain string for headings
        return {"type": _HEADING_ALIAS[typ], "data": str(raw.get("content", "")).strip()}

    if typ == "image":
        typ = "figure"
    elif typ == "formula":
        typ = "equation"

    # ── Format 1: already canonical ───────────────────────────────────────────
    if "data" in raw:
        return {"type": typ, "data": raw["data"]}

    content = raw.get("content")

    # ── Format 2: structured non-text elements using "content" key ────────────
    if typ == "table":
        rows    = content if isinstance(content, list) else []
        style   = raw.get("style", {})
        caption = raw.get("caption", style.get("caption", ""))
        return {"type": "table", "data": {"rows": rows, "caption": caption}}

    if typ == "figure":
        style   = raw.get("style", {})
        caption = raw.get("caption", style.get("caption", ""))
        if isinstance(content, dict):
            # e.g. {"base64": "...", "caption": ""}
            data = dict(content)
            data.setdefault("caption", caption)
        else:
            drawing_xml = raw.get("drawing_xml", style.get("drawing_xml", ""))
            if isinstance(content, str) and content and not drawing_xml:
                drawing_xml = content
            data = {"drawing_xml": drawing_xml, "caption": caption}
        return {"type": "figure", "data": data}

    if typ == "equation":
        style = raw.get("style", {})
        # Collect all equation-related fields that exist on the raw element,
        # falling back to style dict, so parsers producing different schemas
        # (expression/suffix vs omml/ole_base64/image_base64/label) both work.
        data: dict = {}
        for key in ("expression", "suffix", "omml", "ole_base64",
                    "image_base64", "label", "width_pt", "height_pt", "prog_id"):
            val = raw.get(key, style.get(key))
            if val is not None:
                data[key] = val
        # Fallback: bare string content → expression
        if not data and isinstance(content, str):
            data = {"expression": content}
        elif isinstance(content, dict):
            data.update({k: v for k, v in content.items() if k not in data})
        return {"type": "equation", "data": data}

    if typ == "reference":
        style = raw.get("style", {})
        if isinstance(content, dict):
            data = content
        else:
            data = {
                "index":   raw.get("index",   style.get("index", 0)),
                "content": raw.get("ref_content", style.get("content", str(content or ""))),
                "before":  raw.get("before",  style.get("before", "")),
                "after":   raw.get("after",   style.get("after", "")),
            }
        return {"type": "reference", "data": data}

    if typ in ("abstract_cn", "abstract_en"):
        if isinstance(content, dict):
            data = content
        else:
            data = {
                "content":  str(content or ""),
                "keywords": raw.get("keywords", []),
            }
        return {"type": typ, "data": data}

    # ── Format 3: markdown-style body/content (heading prefix or plain text) ──
    text = content if isinstance(content, str) else ""
    m = _MD_HEADING.match(text)
    if m:
        return {"type": f"h{len(m.group(1))}", "data": m.group(2).strip()}
    return {"type": "text", "data": text}


def _load_elements(path: Path) -> list[dict]:
    try:
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)
        items = raw if isinstance(raw, list) else (raw.get("elements") or raw.get("document") or [])
        _normalize_list = []
        for item in items:
            _normalize_list.append(_normalize(item))
        return _normalize_list
    except Exception as e:
        raise e

# ── Front-matter extraction ────────────────────────────────────────────────────

def _dismiss_before_front(elements: list[dict], fm_name_set: set[str]) -> list[dict]:
    """Drop every element that appears before the first front-matter h1 heading."""
    for i, e in enumerate(elements):
        if e["type"] == "h1" and e["data"].strip() in fm_name_set:
            return elements[i:]
    return elements


def _split_front_body(elements: list[dict], fm_name_set: set[str]) -> tuple[list[dict], list[dict]]:
    """
    Split elements into (front_matter_elements, body_elements).

    Front matter ends just before the first h1 heading that is NOT in fm_name_set.
    """
    cut = len(elements)
    for i, e in enumerate(elements):
        if e["type"] == "h1" and e["data"].strip() not in fm_name_set:
            cut = i
            break
    return elements[:cut], elements[cut:]


def _extract_abstract(
    front_elements: list[dict],
    fm_names: list[str],
) -> tuple[str, str, list[str], list[str]]:
    """
    Parse front-matter elements to collect CN/EN abstract content and keywords.

    Returns: (cn_content, en_content, cn_keywords, en_keywords)
    """
    cn_parts:    list[str] = []
    en_parts:    list[str] = []
    cn_keywords: list[str] = []
    en_keywords: list[str] = []

    # Normalise front-matter names for comparison
    name0 = fm_names[0].strip() if len(fm_names) > 0 else None  # 摘  要
    name1 = fm_names[1].strip() if len(fm_names) > 1 else None  # Abstract

    mode = None  # "cn" | "en" | None

    for e in front_elements:
        if e["type"] == "h1":
            name = e["data"].strip()
            if name == name0:
                mode = "cn"
            elif name == name1:
                mode = "en"
            else:
                mode = None          # 目录 or other: skip heading, stop accumulating
            continue                 # never emit front-matter h1 headings

        if mode == "cn" and e["type"] == "text":
            t = e["data"].strip()
            if _KW_CN_RE.match(t):
                body = _KW_CN_RE.sub("", t)
                cn_keywords = [k.strip() for k in re.split(r"[；;]", body) if k.strip()]
            elif t:
                cn_parts.append(t)

        elif mode == "en" and e["type"] == "text":
            t = e["data"].strip()
            if _KW_EN_RE.match(t):
                body = _KW_EN_RE.sub("", t)
                en_keywords = [k.strip() for k in re.split(r"[,，]", body) if k.strip()]
            elif t:
                en_parts.append(t)
        # Non-text elements inside front matter are silently ignored

    return " ".join(cn_parts), " ".join(en_parts), cn_keywords, en_keywords


# ── Action builders ───────────────────────────────────────────────────────────

def _sb_action(sb: dict) -> dict:
    """Build an insert_section_break action from a section_break config dict."""
    a: dict = {
        "type":        "insert_section_break",
        "header_refs": sb.get("header_refs", {}),
        "footer_refs":  sb.get("footer_refs",  {}),
    }
    if "restart_page_number" in sb:
        a["restart_page_number"] = sb["restart_page_number"]
    return a


# ── Main generator ─────────────────────────────────────────────────────────────

def generate(
    full_parsed_path: str = str(_FULL_PARSED),
    config_path:      str = str(_HIT_CONFIG),
    output_path:      str = str(_OUTPUT),
) -> None:
    # ── Load inputs ───────────────────────────────────────────────────────────
    src = Path(full_parsed_path)
    if not src.exists():
        print(f"[ERROR] full_parsed not found: {src}", file=sys.stderr)
        sys.exit(1)

    elements = _load_elements(src)

    with open(Path(config_path), encoding="utf-8") as f:
        cfg: dict = json.load(f)

    front_matter = cfg.get("front_matter", [])
    fm_names     = [fm["name"] for fm in front_matter]
    fm_name_set  = {n.strip() for n in fm_names}

    dismiss        = cfg.get("dismiss_before_front", True)
    toc_level      = cfg.get("toc_max_level", 4)
    bssb           = cfg.get("body_start_section_break")
    body_sbs       = list(cfg.get("body_section_breaks", []))
    final_sb       = cfg.get("final_section_break")
    pb_before_h1   = cfg.get("page_break_before_body_h1", False)

    # ── 1. Dismiss before front ───────────────────────────────────────────────
    if dismiss and fm_name_set:
        elements = _dismiss_before_front(elements, fm_name_set)

    # ── 2. Check for explicit abstract_cn / abstract_en elements ─────────────
    explicit_cn: dict | None = None
    explicit_en: dict | None = None
    cleaned: list[dict] = []
    for e in elements:
        if e["type"] == "abstract_cn" and explicit_cn is None:
            explicit_cn = e["data"] if isinstance(e["data"], dict) else {"content": str(e["data"])}
        elif e["type"] == "abstract_en" and explicit_en is None:
            explicit_en = e["data"] if isinstance(e["data"], dict) else {"content": str(e["data"])}
        else:
            cleaned.append(e)
    elements = cleaned

    # ── 3. Split front matter / body ─────────────────────────────────────────
    front_elems, body_elems = _split_front_body(elements, fm_name_set)

    # ── 4. Extract abstract from front elements (if not already explicit) ─────
    if explicit_cn or explicit_en:
        cn_content  = (explicit_cn or {}).get("content", "")
        en_content  = (explicit_en or {}).get("content", "")
        cn_keywords = (explicit_cn or {}).get("keywords", [])
        en_keywords = (explicit_en or {}).get("keywords", [])
    else:
        cn_content, en_content, cn_keywords, en_keywords = _extract_abstract(
            front_elems, fm_names
        )

    # ── 5. Build actions ──────────────────────────────────────────────────────
    actions: list[dict] = []

    # 5a. Decide which front_matter sections are active based on detected data:
    #       CN abstract section  → only when cn_content was found
    #       EN abstract section  → only when en_content was found
    #       auto_generate_toc    → always (structural requirement, data-independent)
    #       any other entry      → always
    name0 = fm_names[0].strip() if len(fm_names) > 0 else None
    name1 = fm_names[1].strip() if len(fm_names) > 1 else None

    active_fm: list[dict] = []
    for fm in front_matter:
        nm = fm["name"].strip()
        if fm.get("auto_generate_toc"):
            active_fm.append(fm)
        elif nm == name0:
            if cn_content:
                active_fm.append(fm)
        elif nm == name1:
            if en_content:
                active_fm.append(fm)
        else:
            active_fm.append(fm)

    # 5b. Abstract block (only when at least one side has content)
    if cn_content or en_content:
        actions.append({
            "type":             "insert_abstract_with_keywords",
            "cn_content":       cn_content,
            "en_content":       en_content,
            "cn_keywords":      cn_keywords,
            "en_keywords":      en_keywords,
            "cn_title":         fm_names[0] if fm_names else "摘  要",
            "en_title":         fm_names[1] if len(fm_names) > 1 else "Abstract",
            "keyword_label_cn": "关键词",
            "keyword_label_en": "Keywords",
        })

    # 5c. Active front-matter section breaks; auto_generate_toc also emits TOC
    # Only emit front-matter structure when there is actual abstract/front content.
    has_front_matter = bool(cn_content or en_content)
    if has_front_matter:
        for fm in active_fm:
            sb = fm.get("section_break", {})
            if sb:
                actions.append(_sb_action(sb))
            if fm.get("auto_generate_toc"):
                actions.append({"type": "generate_toc", "max_level": toc_level})

    # 5d. Body-start section break (ends TOC section, restarts page at 1)
    if bssb and has_front_matter:
        actions.append(_sb_action(bssb))

    # 5d. Body elements
    body_sb_idx   = 0
    first_h1_seen = False

    for elem in body_elems:
        t    = elem["type"]
        data = elem["data"]

        if t == "h1":
            # Every body h1 after the first needs a page/section break before it.
            if first_h1_seen:
                if body_sb_idx < len(body_sbs):
                    # Consume next section break from pool (sectPr defaults to
                    # nextPage, so this also forces a new page).
                    actions.append(_sb_action(body_sbs[body_sb_idx]))
                    body_sb_idx += 1
                elif pb_before_h1:
                    # Pool exhausted but page break is required — use a simple
                    # page break (no header/footer change).
                    actions.append({"type": "insert_page_break"})
            first_h1_seen = True
            actions.append({"type": "add_heading", "text": data, "level": 1})

        elif t in ("h2", "h3", "h4"):
            actions.append({"type": "add_heading", "text": data, "level": int(t[1])})

        elif t == "text":
            if str(data).strip():
                actions.append({"type": "add_paragraph", "text": data, "style_type": "正文"})

        elif t == "figure":
            a: dict = {"type": "insert_figure"}
            if isinstance(data, dict):
                a.update({k: v for k, v in data.items()})
            actions.append(a)

        elif t == "table":
            a = {"type": "insert_table"}
            if isinstance(data, dict):
                a.update({k: v for k, v in data.items()})
            elif isinstance(data, list):
                a["data"] = data
            actions.append(a)

        elif t == "equation":
            a = {"type": "insert_equation"}
            if isinstance(data, dict):
                a.update({k: v for k, v in data.items()})
            elif isinstance(data, str):
                a["expression"] = data
            actions.append(a)

        elif t == "reference":
            a = {"type": "add_reference"}
            if isinstance(data, dict):
                a.update({k: v for k, v in data.items()})
            actions.append(a)

        # Unknown types are silently skipped

    # 5e. Final section break
    if final_sb:
        actions.append(_sb_action(final_sb))

    # 5f. Strip trailing insert_section_break (avoids blank page at end)
    if actions and actions[-1]["type"] == "insert_section_break":
        actions.pop()

    # ── 6. Write output ───────────────────────────────────────────────────────
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"document": actions}, f, ensure_ascii=False, indent=2)

    print(f"[OK] {len(actions)} actions → {out.resolve()}")
    print(f"     front matter sections  : {len(active_fm)} / {len(front_matter)} active")
    print(f"     body section breaks    : {body_sb_idx} / {len(body_sbs)} consumed")
    print(f"     body level-1 headings  : {sum(1 for e in body_elems if e['type'] == 'h1')}")
    print(f"     abstract CN            : {bool(cn_content)} ({len(cn_keywords)} keywords)")
    print(f"     abstract EN            : {bool(en_content)} ({len(en_keywords)} keywords)")


if __name__ == "__main__":
    args = sys.argv[1:]
    generate(
        full_parsed_path = args[0] if len(args) > 0 else str(_FULL_PARSED),
        config_path      = args[1] if len(args) > 1 else str(_HIT_CONFIG),
        output_path      = args[2] if len(args) > 2 else str(_OUTPUT),
    )
