# DOCX Engine V3 — Phase One Technical Reference

> **Purpose**: Complete reference manual for maintaining and modifying
> `docx_extractor.py` and `docx_compiler.py`.  Documents every module's
> design rationale, data model, function responsibilities, critical
> implementation details, and known pitfalls.
>
> **Last updated**: reflects the raw-XML preservation revision described in
> `REVIEW.md` (TOC lossless round-trip via `type: "raw_xml"`).

---

## Table of Contents

1. [Overall Pipeline](#1-overall-pipeline)
2. [extraction.json Data Model](#2-extractionjson-data-model)
3. [docx_extractor.py](#3-docx_extractorpy)
4. [docx_compiler.py](#4-docx_compilerpy)
5. [Key Design Decisions & Known Pitfalls](#5-key-design-decisions--known-pitfalls)
6. [Maintenance Guide: Common Modification Scenarios](#6-maintenance-guide-common-modification-scenarios)

---

## 1. Overall Pipeline

```
template.docx
    │
    ▼  DocxExtractor.extract()
extraction.json          ← structured intermediate representation, human-readable
    │
    ▼  DocxCompiler.compile()   +   template/  (template directory)
output.docx
```

### Role of the template directory (`template/`)

The extractor unzips the `.docx` here; the compiler copies binary assets from
here.  Both tools share the same directory:

```
template/
  word/
    document.xml            ← body XML  (compiler rewrites completely)
    header1.xml …           ← headers   (compiler rewrites completely)
    footer1.xml …           ← footers   (compiler rewrites completely)
    styles.xml              ← styles    (copied unchanged)
    settings.xml            ← settings  (compiler injects updateFields)
    _rels/document.xml.rels ← relations (compiler appends new rIds for images/OLE)
    media/                  ← images    (copied unchanged; new images appended)
    embeddings/             ← OLE bins  (copied unchanged; new bins appended)
  [Content_Types].xml       ← content types (compiler appends new extensions)
  docProps/                 ← metadata  (copied unchanged)
```

> **Core principle**: `extraction.json` is the authoritative source for all
> *content*.  The template directory supplies *structural* data that
> `extraction.json` does not store:
> - Table skeleton XML (`tblPr / tblGrid / trPr / tcPr`)
> - Body-level `sectPr` (page size, margins, full section layout)
> - Binary assets (`media/`, `embeddings/`)
> - `styles.xml`, `theme/`

---

## 2. extraction.json Data Model

The extractor outputs a single JSON object with these top-level fields:

```jsonc
{
  "source":        "template.docx",
  "extracted_at":  "2026-04-06T10:00:00",
  "metadata":      { "Title": "…", … },
  "placeholders":  ["field_a", "field_b"],
  "sections":      [ … ],
  "body_elements": [ … ],
  "headers":       { "header1.xml": { … } },
  "footers":       { "footer1.xml": { … } },
  "relationships": { "rId1": { … }, … }
}
```

---

### 2.1 body_elements list

Each element is a dict with a mandatory `"type"` field.

#### type = "paragraph"

```jsonc
{
  "index":  42,
  "type":   "paragraph",
  "style":  "Heading1",       // w:pStyle val (may be null)
  "text":   "Chapter 1 …",   // full paragraph text (for search/debug only)
  "pPr":    "<w:pPr …/>",    // paragraph-properties XML string (ns0: prefixes)
  "runs":   [ … ],            // list of run dicts (see 2.2)
  "section_break": null       // or { header_refs, footer_refs, page_size }
}
```

#### type = "raw_xml"

Introduced to achieve **lossless round-trip** of Word field structures and
structured document tags.  These elements must never be parsed or regenerated
— doing so breaks TOC hyperlinks, page numbers, and cross-references.

```jsonc
{
  "index": 89,
  "type":  "raw_xml",
  "xml":   "<w:p>…</w:p>"    // complete serialised XML of one body-level element
}
```

**When the extractor emits `raw_xml`**:

| Trigger | Description |
|---------|-------------|
| `w:p` containing `w:fldChar fldCharType="begin"` where the field instruction starts with a preserved keyword | Each paragraph in the field block (from `begin` to `end`) is stored as a separate `raw_xml` entry |
| `w:sdt` (Structured Document Tag) at body level | The entire `sdt` element is stored as one `raw_xml` entry |

**Preserved field keywords** (`_PRESERVED_FIELDS`):

```
TOC  REF  SEQ  PAGE  NUMPAGES  STYLEREF
```

#### type = "table"

```jsonc
{
  "index": 10,
  "type":  "table",
  "rows": [
    [
      { "text": "cell text", "paragraphs": [ {paragraph}, … ] },
      …
    ],
    …
  ]
}
```

Note: table structural XML (`tblPr`, `tblGrid`, `trPr`, `tcPr`) is **not**
stored here — the compiler reads it from the template table at the matching
position.

#### type = "image" (added by DocxTools)

```jsonc
{
  "type":     "image",
  "base64":   "iVBORw0KGgo…",
  "caption":  "Figure 2-1 …",
  "position": "center",
  "width":    120,
  "height":   80
}
```

> Template inline images are **not** extracted as `type:"image"`.  They are
> stored as `drawing_xml` inside their paragraph's run list (see 2.2) and
> are re-emitted verbatim by `_build_para`.

#### type = "toc" (added by DocxTools)

```jsonc
{ "type": "toc", "max_level": 4 }
```

Used only when DocxTools explicitly inserts a new TOC.  Template TOCs that
existed in the original docx are preserved as `raw_xml` elements (above) and
are never regenerated.

#### type = "ole" (added by DocxTools)

```jsonc
{
  "type":          "ole",
  "base64":        "…",
  "formula_index": "(2-1)"
}
```

#### type = "omath" / "omathpara" (OMML math)

```jsonc
{
  "type":          "omath",
  "formula":       "<m:oMath>…</m:oMath>",
  "formula_index": "(3-1)"
}
```

#### Other types (internal markers)

Tags such as `"bookmarkEnd"` and `"bookmarkStart"` are recorded as
`{"index": n, "type": "bookmarkEnd"}`.  The compiler skips these (no content
to reconstruct).

---

### 2.2 Run data structure

```jsonc
{
  "text":        "plain text",
  "rPr":         "<w:rPr …/>",     // character properties XML (may be null)
  "drawing_xml": "<w:drawing …/>", // present for image/shape runs
  "object_xml":  "<w:object …/>"   // present for OLE formula runs
}
```

Each run is exactly one of three mutually exclusive forms:

| Form | Fields present |
|------|---------------|
| Text run | `text` (may be empty string), `rPr` optional |
| Drawing run | `drawing_xml`, `rPr` optional |
| OLE run | `object_xml`, `rPr` optional |

`drawing_xml` may be any of:

1. `<w:drawing>` containing `wp:inline` — inline image (flows with text)
2. `<w:drawing>` containing `wp:anchor` — floating shape (non-WPS, rare)
3. `<mc:AlternateContent>` wrapping `mc:Choice/w:drawing/wp:anchor` — WPS
   floating shape (most common anchor type); **must be stored whole** so WPS
   can read `posOffset` values

---

### 2.3 sections list

```jsonc
{
  "paragraph_index": 42,
  "header_refs": { "default": "rId5", "first": "rId6" },
  "footer_refs": { "default": "rId7" },
  "page_size":   { "w": "12240", "h": "15840" }
}
```

`paragraph_index` is `null` for the final body-level section.

---

### 2.4 headers / footers

```jsonc
{
  "header1.xml": {
    "text":       "full text (debug only)",
    "paragraphs": [ {paragraph}, … ]
  }
}
```

Keys are filenames relative to `word/`.

---

### 2.5 relationships

```jsonc
{
  "rId1": { "type": "header",    "target": "header1.xml" },
  "rId8": { "type": "image",     "target": "media/image1.png" },
  "rId9": { "type": "oleObject", "target": "embeddings/oleObject1.bin" }
}
```

---

## 3. docx_extractor.py

### 3.1 Namespace constants

```python
W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
MC  = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
```

The extractor does **not** call `ET.register_namespace()`, so `ET.tostring()`
produces `ns0:`, `ns1:` auto-prefixes in stored XML strings.  The compiler
resolves these back to standard prefixes via `ET.fromstring()` (which converts
to Clark notation) + `ET.tostring()` (which uses the compiler's registered
prefixes).  No manual string rewriting is needed.

### 3.2 Field preservation constant

```python
_PRESERVED_FIELDS = frozenset({'TOC', 'REF', 'SEQ', 'PAGE', 'NUMPAGES', 'STYLEREF'})
```

Any `w:fldChar` field whose instruction text begins with one of these keywords
is preserved as `raw_xml` rather than parsed.  See §5.1 for the rationale.

---

### 3.3 Module-level utility functions

| Function | Purpose |
|----------|---------|
| `_q(local)` | Returns `{W}local` (Clark notation for `w:` namespace) |
| `_tag(elem)` | Returns local tag name (strips namespace URI) |
| `_text_of(elem)` | Concatenates all `w:t` text under `elem` recursively |
| `_is_inline_drawing(xml_str)` | Checks whether XML string wraps a `wp:inline` element (legacy helper, not called in main path) |

---

### 3.4 Field block detection functions

#### `_has_field_begin(p_elem) → bool`

Returns `True` if `p_elem` (a `<w:p>`) contains any `w:fldChar` with
`fldCharType="begin"`.  Used as a fast pre-check before attempting block
collection.

#### `_field_depth_delta(elem) → int`

Iterates all `w:fldChar` descendants of `elem` and computes the net change in
field nesting depth: each `begin` contributes +1, each `end` contributes -1.
`separate` does not affect depth.

#### `_collect_field_block(children, start) → (block | None, consumed)`

Starting at `children[start]` (a paragraph that passed `_has_field_begin`),
collects body-level elements until the outermost field is closed:

```python
depth = 0
while j < len(children):
    block.append(children[j])
    depth += _field_depth_delta(children[j])
    j += 1
    if depth <= 0:
        break
```

After collection, gathers all `instrText` content from the block.  If the
first whitespace-delimited token of the instruction (upper-cased) is in
`_PRESERVED_FIELDS`, returns `(block, consumed)`.  Otherwise returns
`(None, 0)` so the caller falls through to normal paragraph extraction.

**Multi-paragraph field handling**: A TOC typically spans dozens of paragraphs
(begin paragraph → many entry paragraphs → end paragraph).  This function
collects all of them in a single pass using depth tracking, regardless of how
many paragraphs the field spans.

---

### 3.5 Run extraction: `_extract_run(r_elem)`

**Input**: a single `<w:r>` element.

**Priority order**:

```
1. Check direct children for w:drawing
   ├─ Not found? Look for mc:AlternateContent child
   │            with mc:Choice containing w:drawing
   │            → store complete mc:AlternateContent as drawing_xml, return
   └─ Found w:drawing → store as drawing_xml, return

2. Found w:object? → store as object_xml, return

3. Default: read w:t text → store as text
```

**Why store the complete `mc:AlternateContent`**: WPS Office wraps its
proprietary vector shapes (text boxes, connectors, arrows) in
`mc:AlternateContent`.  WPS needs the wrapper intact to read the
`posOffset` values (absolute EMU position).  Storing only the inner
`w:drawing` causes WPS to render the shape at a random position.

---

### 3.6 Paragraph extraction: `_extract_paragraph(p_elem, index)`

| Field | Source | Compiler use |
|-------|--------|-------------|
| `style` | `w:pPr/w:pStyle[@w:val]` | Debug only; compiler uses `pPr` string |
| `text` | `_text_of(p_elem)` | Debug / placeholder search only |
| `runs` | All direct `<w:r>` children | Content reconstruction |
| `pPr` | Complete `<w:pPr>` XML string | Paragraph formatting (indent, spacing, etc.) |
| `section_break` | `w:pPr/w:sectPr` | Non-final section metadata |

> **Known limitation**: only direct `<w:r>` children are collected.  Runs
> inside `w:hyperlink`, `w:ins` (tracked insertions), or other wrappers are
> skipped.  If the template uses revision tracking or hyperlinks, those runs
> will not appear in extraction.

---

### 3.7 Table extraction

Call chain: `_extract_table` → `_extract_cell` → `_extract_paragraph`

Only content is stored (cell paragraphs with their text and formatting).
Structural XML (`tblPr`, `tblGrid`, `trPr`, `tcPr`) is **not stored** —
the compiler reads it from the corresponding template table.

---

### 3.8 Header / footer extraction: `_extract_hf_file(xml_path)`

Reads a single `headerN.xml` or `footerN.xml`.  Uses `root.iter(_q('p'))`
(depth-first) so paragraphs nested inside tables are also captured.

**Current limitation**: headers/footers containing field structures (`PAGE`,
`NUMPAGES`, `STYLEREF`) are extracted paragraph-by-paragraph via
`_extract_paragraph`.  Runs that contain `w:fldChar` or `w:instrText` (with
no `w:t`) are extracted with `text = ""`, losing page-number field logic.
Field-block preservation is applied only to the document body, not to
header/footer XML.  This is acceptable for most templates where header/footer
fields are static or regenerated by Word on open.

---

### 3.9 `DocxExtractor.extract()` main flow

```
1. Unzip .docx → unzip_dir
   (skipped if directory already exists)

2. Parse word/_rels/document.xml.rels → relationships dict

3. Iterate body direct children (index-based while loop):

   child tag == "p" AND has fldChar begin?
   ├─ _collect_field_block()
   │  ├─ Field type in _PRESERVED_FIELDS?
   │  │  → emit one raw_xml entry per element in block
   │  │    advance i by consumed, continue
   │  └─ Not preserved → fall through to normal paragraph handling

   child tag == "p"         → _extract_paragraph() → paragraph entry
                               if section_break → append to sections
   child tag == "sdt"       → raw_xml entry (entire sdt element)
   child tag == "tbl"       → _extract_table() → table entry
   child tag == "sectPr"    → _extract_sectPr() → sections (paragraph_index=None)
   other tags               → {"index": n, "type": "<tag>"} (internal marker)

4. Extract header/footer files referenced in relationships

5. Scan all paragraph text for {{placeholder}} patterns

6. Load metadata from docProps/

7. Serialize to data/extraction.json
```

---

## 4. docx_compiler.py

### 4.1 Namespace registration

```python
for _p, _u in _NS_MAP.items():
    ET.register_namespace(_p, _u)
```

**Must execute before any `ET.parse()` or `ET.tostring()` call** — placed at
module top level.  Registered namespaces ensure `ET.tostring()` outputs
standard prefixes (`w:`, `wp:`, `a:`, `mc:`, etc.) rather than auto-generated
`ns0:` prefixes.  This is critical for WPS's XML parser, which rejects
non-standard prefixes in DrawingML elements.

The extractor's `ns0:` prefixes in stored `pPr`/`rPr` strings are not a
problem: `ET.fromstring()` converts them to Clark notation, and
`ET.tostring()` re-serialises with the registered prefixes.

---

### 4.2 Namespace URI constants

```python
W   = '…/wordprocessingml/2006/main'
M   = '…/officeDocument/2006/math'
R   = '…/officeDocument/2006/relationships'
REL = '…/package/2006/relationships'
A   = '…/drawingml/2006/main'
PIC = '…/drawingml/2006/picture'
WP  = '…/drawingml/2006/wordprocessingDrawing'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
V_NS = 'urn:schemas-microsoft-com:vml'
O_NS = 'urn:schemas-microsoft-com:office:office'
MC   = '…/markup-compatibility/2006'
```

---

### 4.3 Module-level utility functions

| Function | Notes |
|----------|-------|
| `_q(local)` | `{W}local` |
| `_qm(local)` | `{M}local` (OMML math namespace) |
| `_qr(local)` | `{R}local` (relationship namespace) |
| `_tag(elem)` | Local tag name |
| `_parse_xml(xml_str)` | Parse XML string → Element; returns `None` on error |
| `_pack_docx(source_dir, output_path)` | Zip directory as .docx (sorted filenames, deterministic) |
| `_write_xml(path, root)` | Write Element to file with `<?xml …?>` declaration |
| `_image_suffix(data)` | Detect image format from magic bytes; returns `.png`, `.jpg`, etc. |
| `_rels_append(rels_path, rid, rel_type, target)` | Append a `<Relationship>` using string replacement (avoids default-namespace loss) |
| `_ensure_content_type(ct_path, ext)` | Add `<Default>` content-type entry if absent (string replacement) |
| `_minimal_tbl()` | Bare valid `<w:tbl>` with `TableGrid` style |
| `_minimal_tc()` | Bare valid `<w:tc>` with empty `<w:tcPr>` |
| `_plain_para(text)` | Unstyled text paragraph (fallback / empty cell filler) |
| `_math_text(oMath, text)` | Append plain-text `<m:r><m:t>` run to an `<m:oMath>` |

> **Why `_rels_append` and `_ensure_content_type` use string operations**:
> `.rels` files and `[Content_Types].xml` use default namespaces
> (`xmlns="…"` with no prefix).  ET re-serialisation loses the default
> namespace declaration, corrupting those files.  String-based insertion is
> safe and unambiguous.

---

### 4.4 WPS anchor helpers

#### `_WPS_URIS` (set)

GraphicData URIs that require `mc:AlternateContent` wrapping:

```python
_WPS_URIS = {
    '…/wordprocessingShape',
    '…/wordprocessingGroup',
    '…/wordprocessingInk',
    '…/wordprocessingCanvas',
}
```

#### `_is_wps_anchor_drawing(elem) → bool`

Returns `True` when `elem` is a bare `<w:drawing>` containing a `wp:anchor`
whose `a:graphicData[@uri]` is in `_WPS_URIS`.

#### `_wrap_in_mc_choice(drawing_elem) → ET.Element`

Wraps a bare `<w:drawing>` in `<mc:AlternateContent><mc:Choice Requires="wps">`.
Used as a backward-compatibility shim for extraction data produced before the
extractor was updated to store the full `mc:AlternateContent`.  Current
extraction stores the complete wrapper, so this function is only a safety net.

---

### 4.5 `DocxCompiler.__init__`

```python
DocxCompiler(
    extraction_path = 'data/extraction.json',
    template_dir    = 'template',
)
```

- Validates both paths exist.
- Loads `extraction.json` into `self.ext`.
- `self._rid_counter`: reset by `_init_rid_counter` before compile.
- `self._shape_counter`: `docPr` id for inline images / OLE shape IDs.

---

### 4.6 `compile()` main flow

```
1. Copy template/ → work_dir  (original never modified)

2. _init_rid_counter(work_dir)
   → parse document.xml.rels, set _rid_counter = max(existing) + 1

3. _rebuild_document(work_dir)
   → completely rewrite word/document.xml

4. _rebuild_hf_files(work_dir)
   → completely rewrite all header/footer XML files

5. _patch_settings(work_dir)
   → inject <w:updateFields w:val="true"/> into settings.xml

6. _pack_docx(work_dir, output_path)
   → zip work_dir as the final .docx

7. Cleanup temporary directory (finally block)
```

---

### 4.7 `_rebuild_document(work_dir)`

**Purpose**: Completely replace `word/document.xml` body content.

```
1. Parse template document.xml
   ├─ Collect template tables: tmpl_tables (k-th used for k-th extraction table)
   └─ Deep-copy body-level <w:sectPr> before clearing

2. Clear all children from <w:body>

3. Iterate extraction body_elements, dispatch by type:

   "paragraph"         → _build_para()
   "raw_xml"           → _parse_xml(elem["xml"]) → append directly to body
   "toc"               → _build_toc_para()
   "table"             → _build_table()  (k-th template table as skeleton)
   "image"             → _build_image_nodes()  (returns 1-2 paragraphs)
   "omath"/"omathpara" → _build_omath_para()
   "ole"               → _build_ole_para()
   other types         → skipped (bookmarkEnd, etc. — no content to restore)

4. Append deep-copied body-level <w:sectPr> last
   (preserves page size, margins, header distance, etc.)

5. _write_xml(doc_path, root)
```

**raw_xml dispatch** (critical — added per `REVIEW.md`):

```python
elif etype == 'raw_xml':
    node = _parse_xml(elem.get('xml', ''))
    if node is not None:
        body.append(node)
    stats['raw_xml'] += 1
```

The stored XML string is a single serialised element (one `<w:p>` or one
`<w:sdt>`).  `_parse_xml` calls `ET.fromstring()` which resolves any `ns0:`
prefixes to Clark notation; `ET.tostring()` (called internally by
`_write_xml`) re-serialises with the compiler's registered standard prefixes.
The result is semantically identical to the original template XML.

---

### 4.8 `_build_para(pdata) → ET.Element`

Reconstructs a `<w:p>` from an extraction paragraph dict.

```xml
<w:p>
  [<w:pPr>…</w:pPr>]       ← parsed from pdata['pPr'] string
  <w:r>                      ← one per run in pdata['runs']
    [<w:rPr>…</w:rPr>]
    (one of three content forms)
  </w:r>
</w:p>
```

**Run content decision tree**:

```
run has drawing_xml?
├─ root tag == AlternateContent → append directly (WPS shape with mc:Fallback)
└─ root tag == drawing
   ├─ _is_wps_anchor_drawing? → _wrap_in_mc_choice() then append (backward compat)
   └─ else (wp:inline image)  → append directly

run has object_xml?
└─ parse and append directly (OLE formula, re-emitted verbatim)

otherwise (text run)
└─ create <w:t>, set xml:space="preserve" if leading/trailing whitespace
```

---

### 4.9 `_build_toc_para(elem) → ET.Element`

Generates a fresh TOC field paragraph for **DocxTools-added** TOCs only.
Template TOCs are preserved as `raw_xml` and never reach this function.

```xml
<w:p>
  <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
  <w:r><w:instrText xml:space="preserve"> TOC \o "1-N" \z </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
```

| Switch | Meaning |
|--------|---------|
| `\o "1-N"` | Include outline heading levels 1 through `max_level` |
| `\z` | Hide tab leader and page numbers in Web Layout view |
| `w:dirty="true"` | Cached content is stale; Word/WPS will regenerate on open |
| `\h` | **Intentionally omitted**: would produce blue hyperlink styling — not standard for Chinese academic thesis TOCs |

---

### 4.10 `_build_table(tdata, tmpl_tbl) → ET.Element`

**Content source**: extraction.json cell paragraphs.
**Structure source**: template table (`tblPr`, `tblGrid`, `trPr`, `tcPr`).

**Row/column fallback rules**:

| Situation | Handling |
|-----------|----------|
| Row index within template row count | Use that template row |
| Row index exceeds template | Clone last template row |
| Column index within template column count | Use that template cell |
| Column index exceeds template | Clone last template cell |
| No template table at all | Use `_minimal_tbl()` / `_minimal_tc()` |

Each cell's `tcPr` (column width, borders, merge markers) is inherited from
the template clone; paragraph content is rebuilt from extraction via
`_build_para()`.

---

### 4.11 `_build_image_nodes(elem, work_dir) → list[ET.Element]`

Returns one or two `<w:p>` elements (image + optional caption).

#### Path A — `drawing_xml` (template-preserved image)

Parse and re-emit the stored XML directly inside a `<w:r>`.  The rId remains
valid because `word/media/` and `word/_rels/` are copied from the template.

#### Path B — `base64` (DocxTools-added image)

```
1. Decode base64 bytes
2. Detect format suffix via magic bytes (_image_suffix)
3. Write to word/media/imageN.ext
4. Append new Relationship to document.xml.rels (_rels_append)
5. Ensure [Content_Types].xml has the extension (_ensure_content_type)
6. Build DrawingML wp:inline paragraph (_drawing_para)
```

Falls back to a `[图片]` plain-text placeholder when base64 decoding fails.

---

### 4.12 `_build_ole_para(elem, work_dir) → ET.Element`

Builds a centred paragraph containing an OLE embedded object (Equation
Editor formula):

```xml
<w:p>
  <w:pPr><w:jc w:val="center"/></w:pPr>
  <w:r>
    <w:object w:dxaOrig="2400" w:dyaOrig="600">
      <v:shape id="_x0000_i…" style="width:120pt;height:30pt" o:ole=""/>
      <o:OLEObject Type="Embed" ProgID="Equation.3"
                   r:id="rIdN" ShapeID="…" DrawAspect="Content"/>
    </w:object>
  </w:r>
  [<w:r><w:t>  (2-1)</w:t></w:r>]   ← formula_index, if present
</w:p>
```

No `<v:imagedata>` preview is included; Word/WPS renders live OLE content.

---

### 4.13 `_build_omath_para(elem) → ET.Element`

Builds an OMML math formula paragraph:

```xml
<w:p>
  <m:oMathPara>
    <m:oMath>
      [XML fragment | <m:r><m:t>plain text</m:t></m:r>]
    </m:oMath>
  </m:oMathPara>
  [<w:r><w:t>  (3-1)</w:t></w:r>]   ← formula_index, if present
</w:p>
```

If `formula` begins with `<`, it is parsed as XML; otherwise it is wrapped as
plain text.

---

### 4.14 `_patch_settings(work_dir)`

Injects into `word/settings.xml`:

```xml
<w:updateFields w:val="true"/>
```

Inserted immediately before `</w:settings>` using string replacement.

**Effect**: Word/WPS recalculates all fields (including the TOC) every time
the document is opened.
**Idempotent**: skipped if `updateFields` is already present.

---

### 4.15 `_rebuild_hf_files()` / `_rebuild_hf(xml_path, hf_data)`

**`_rebuild_hf_files`**: merges `self.ext['headers']` and `self.ext['footers']`,
calls `_rebuild_hf` for each referenced file that exists in `work_dir`.

**`_rebuild_hf(xml_path, hf_data)`**:

```
1. Parse template headerN.xml / footerN.xml
   → keep root element <w:hdr> / <w:ftr> (preserves namespace declarations)

2. Remove all children

3. For each paragraph in hf_data['paragraphs']:
   → _build_para()

4. If no paragraphs: append empty <w:p>
   (OOXML requires at least one paragraph per hdr/ftr)

5. _write_xml(xml_path, root)
```

---

## 5. Key Design Decisions & Known Pitfalls

### 5.1 Raw XML preservation for field structures (most important)

**Decision** (from `REVIEW.md`): Word field structures (`w:fldChar` begin →
instrText → separate → content → end) must **never** be parsed and
regenerated.  They represent dynamic content managed by Word's field engine;
manual reconstruction always produces breakage:

- TOC hyperlinks become broken or disappear
- Page numbers show wrong values
- Cross-references (`REF`, `SEQ`) lose their targets

**Implementation**:
- Extractor: body loop uses `_collect_field_block()` to detect and absorb
  entire multi-paragraph field blocks; each paragraph is stored as `raw_xml`
- Compiler: `_rebuild_document` emits `raw_xml` elements via `_parse_xml()`
  + `body.append()` — no structural modification whatsoever

**Also preserved as `raw_xml`**:
- `w:sdt` (Structured Document Tags) — breaking these causes incorrect formula
  numbering and content-control corruption
- Any body-level element whose tag is not `p`, `tbl`, or `sectPr` is stored as
  a placeholder `{"type": "<tag>"}` and silently skipped by the compiler;
  for elements that carry content (like `sdt`) the `raw_xml` path is used

---

### 5.2 Paragraph structure is never split

**Decision**: Every `<w:p>` is extracted as a single `paragraph` element
containing all its runs — inline images, anchor shapes, OLE objects, and text
runs alike.

**Historical lesson**: Early versions promoted inline images (`wp:inline`) to
top-level `type:"image"` elements, removing them from their paragraph.  This
caused:
- Side-by-side images (e.g. (a)(b) comparison figures) to stack vertically
- Anchor shapes positioned `relativeFrom="paragraph"` to float to wrong
  positions after their host paragraph was deleted

**Rule**: **Never extract an image run out of its owning paragraph.**

---

### 5.3 `mc:AlternateContent` completeness

WPS Office wraps all WPS-proprietary vector graphics in `mc:AlternateContent`
including an `mc:Fallback` VML section.  The entire wrapper must be preserved.

- Extractor: `_extract_run` checks for `mc:AlternateContent` before `w:drawing`
- Compiler: `_build_para` checks for `AlternateContent` root tag first;
  for legacy bare `w:drawing` data uses `_wrap_in_mc_choice()` as a safety net

---

### 5.4 Table structure comes from the template

`extraction.json` does not store `tblPr`, `tblGrid`, `trPr`, or `tcPr`.
The compiler reads these from the k-th template table (positional mapping).

**Consequence**: table order must remain consistent between extraction and
compilation.  DocxTools-added tables that have no template counterpart receive
a minimal skeleton with no borders or column widths.

To preserve table structure in extraction: modify `_extract_table()` to add
`tblPr_xml` / `tblGrid_xml` fields, and update `_build_table()` to prefer
these over the template skeleton (with template as fallback).

---

### 5.5 `sectPr` completeness

The body-level `<w:sectPr>` holds dozens of attributes (margins, line
numbers, header distance, page numbering start value, etc.).  `extraction.json`
`sections` stores only three fields from it (`header_refs`, `footer_refs`,
`page_size`).

**Solution**: The compiler saves a `copy.deepcopy` of the template's
`sectPr` before clearing the body, then appends it unchanged at the very end.
The `sections` data in `extraction.json` is used only for informational
purposes; the compiler never attempts to reconstruct `sectPr` from it.

---

### 5.6 Namespace prefix stability

`ET.register_namespace()` must run at module load time (file top level).  If
called after any `ET.parse()`, elements parsed before registration will still
serialise with `ns0:` prefixes in that run, breaking WPS's XML parser.

All registered prefixes — particularly `mc:`, `wp:`, `a:`, `pic:`, `wps:`,
`w14:` — must be registered in `docx_compiler.py` before any XML parsing
begins.

---

### 5.7 `.rels` and `[Content_Types].xml` use string insertion

Both files use XML default namespaces (`xmlns="…"` with no prefix).
Parsing them with ET and re-serialising loses the default namespace declaration,
producing invalid files that cause Word/WPS to refuse to open the document.

`_rels_append` and `_ensure_content_type` therefore use `str.replace()` to
insert new entries immediately before the closing tag.  This is safe as long
as the closing tags (`</Relationships>`, `</Types>`) appear exactly once in the
file (which OOXML guarantees).

---

## 6. Maintenance Guide: Common Modification Scenarios

### Scenario A: Add a new preserved field type

In `docx_extractor.py`, add the keyword to `_PRESERVED_FIELDS`:

```python
_PRESERVED_FIELDS = frozenset({
    'TOC', 'REF', 'SEQ', 'PAGE', 'NUMPAGES', 'STYLEREF',
    'HYPERLINK',   # ← new
})
```

No compiler change needed — all `raw_xml` elements are already emitted
verbatim.

---

### Scenario B: Preserve header/footer field structures

Currently `_extract_hf_file` parses all paragraphs individually, losing
`fldChar`-based field runs (page numbers, section names).  To preserve them,
apply field-block detection in the header/footer loop as well:

```python
def _extract_hf_file(xml_path):
    root = ET.parse(xml_path).getroot()
    children = list(root)
    paras = []
    i = 0
    while i < len(children):
        child = children[i]
        if _tag(child) == 'p' and _has_field_begin(child):
            block, consumed = _collect_field_block(children, i)
            if block is not None:
                for blk_elem in block:
                    paras.append({
                        'index': len(paras),
                        'type':  'raw_xml',
                        'xml':   ET.tostring(blk_elem, encoding='unicode'),
                    })
                i += consumed
                continue
        paras.append(_extract_paragraph(child, len(paras)))
        i += 1
    return {'text': _text_of(root), 'paragraphs': paras}
```

The compiler's `_rebuild_hf` must then handle `raw_xml` paragraphs:

```python
for pdata in hf_data.get('paragraphs', []):
    if pdata.get('type') == 'raw_xml':
        node = _parse_xml(pdata['xml'])
        if node is not None:
            root.append(node)
    else:
        root.append(self._build_para(pdata))
```

---

### Scenario C: Extract table structure into extraction.json

Modify `_extract_table()`:

```python
def _extract_table(tbl, index):
    tblPr   = tbl.find(_q('tblPr'))
    tblGrid = tbl.find(_q('tblGrid'))
    return {
        'index':       index,
        'type':        'table',
        'tblPr_xml':   ET.tostring(tblPr,   encoding='unicode') if tblPr   else None,
        'tblGrid_xml': ET.tostring(tblGrid,  encoding='unicode') if tblGrid else None,
        'rows': [ … ],
    }
```

In `_build_table()`, prefer extraction data over template skeleton:

```python
tblPr_xml = tdata.get('tblPr_xml')
if tblPr_xml and tmpl_tbl is not None:
    tbl = copy.deepcopy(tmpl_tbl)
    old = tbl.find(_q('tblPr'))
    if old is not None:
        tbl.remove(old)
    tbl.insert(0, _parse_xml(tblPr_xml))
```

---

### Scenario D: Support `w:hyperlink` runs in paragraphs

In `_extract_paragraph()`, extend run collection to recurse into hyperlinks:

```python
runs = []
for child in p_elem:
    tag = _tag(child)
    if tag == 'r':
        runs.append(_extract_run(child))
    elif tag == 'hyperlink':
        for r in child.findall(_q('r')):
            runs.append(_extract_run(r))
```

Note: the hyperlink `rId` and `w:hyperlink` wrapper element are lost — only
run text and formatting survive.  For full hyperlink round-trip, store the
`w:hyperlink` element itself as `raw_xml` instead.

---

### Scenario E: Modify TOC switches (DocxTools-added TOCs only)

In `DocxCompiler._build_toc_para()`, edit the instruction string.  Examples:

```python
# Restore hyperlink entries (blue underlined — non-standard for HIT thesis)
instr = f' TOC \\o "1-{max_level}" \\h \\z '

# Include all heading levels without depth limit
instr = ' TOC \\h \\z '
```

Template TOCs (stored as `raw_xml`) are unaffected by this change.

---

### Scenario F: Auto-detect image pixel dimensions

In `_build_image_nodes()`, Path B, after decoding the base64 bytes:

```python
from PIL import Image
from io import BytesIO

with Image.open(BytesIO(img_bytes)) as im:
    px_w, px_h = im.size
    dpi_x, dpi_y = im.info.get('dpi', (96, 96))
    cx = int(px_w / dpi_x * 914400)   # pixels → EMU
    cy = int(px_h / dpi_y * 914400)
```

Replace the `width_pt` / `height_pt` → EMU conversion block with the above.
