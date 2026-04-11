# Project Checkpoint

**Date:** 2026-04-05  
**Stage:** Core engine complete — extraction and lossless restoration working

---

## Completed Work

### Files in place

| File | Status | Purpose |
|---|---|---|
| `template.docx` | Source | Original 18-page HIT thesis template (WPS Office) |
| `template/` | Extracted | Unzipped OOXML package — never modified by the engine |
| `docx_extractor.py` | Complete | Parses template → `data/extraction.json` |
| `docx_restore.py` | Complete | Fills `{{placeholders}}` → `output.docx` |
| `data/extraction.json` | Generated | 633 KB structured dump of the full template |
| `DESIGN.md` | Complete | Technical writeup of lossless round-trip design |
| `CLAUDE.md` | Complete | Repo guidance for future Claude sessions |

### What the extractor produces (`data/extraction.json`)

```
body_elements : 266  (261 paragraphs, 4 tables, 1 bookmarkEnd)
sections      : 11
headers       : 20   (header1.xml … header20.xml)
footers       : 16   (footer1.xml … footer16.xml)
relationships : 58   (20 headers, 16 footers, 11 images, 6 OLE, styles, theme, …)
placeholders  : []   (none yet — template not yet annotated)
```

### What the restorer does

- Copies `template/` to a `tempfile` working directory (original never touched)
- Applies `{{key}} → value` substitutions to `document.xml` + all headers/footers
- Handles split-run placeholders (merges `w:r` text, rebuilds into first run)
- Preserves all `w:rPr`, `w:pPr`, `w:sectPr`, images, OLE objects verbatim
- Repacks as a valid `.docx` ZIP with deterministic file order
- Verified working via round-trip test

---

## Template Structure Reference

### Document sections (11 total)

| # | Break at para | Headers present | Footers present | Corresponds to |
|---|---|---|---|---|
| 1 | 18 | first, default, even | default, even | Cover page |
| 2 | 62 | — | default, even | 说明页 (instructions) |
| 3 | 75 | default, even | default, even | 中文摘要 |
| 4 | 87 | default, even | default | 英文 Abstract |
| 5 | 117 | default, even | even | 目录 (TOC) |
| 6 | 169 | default, even | default, even | 正文 chapters (main body) |
| 7 | 222 | default | default | 结论 (conclusions) |
| 8 | 234 | default, even | default, even | 参考文献 (references) |
| 9 | 246 | default, even | — | 攻读学位期间成果 |
| 10 | 260 | default, even | default, even | 致谢 (acknowledgements) |
| 11 | body | default, even | default | Final / trailing section |

All sections use A4 page size (`w=11906`, `h=16838` in twentieths of a point).

### Tables (4 total)

| Body index | Dimensions | Content |
|---|---|---|
| 58 | 7 × 3 | Cover page info table (student name, ID, supervisor, major, etc.) |
| 190 | 8 × 6 | Data table — measurement results set 1 |
| 206 | 5 × 6 | Data table — measurement results set 2 |
| 215 | 4 × 6 | Data table — measurement results set 3 (continuation) |

The cover page table (index 58) is the primary target for placeholder injection.

### Binary assets (never modified)

| Type | Count | Location |
|---|---|---|
| TIFF images | 5 | `word/media/image1–5.tiff` |
| WMF images | 6 | `word/media/image6–11.wmf` |
| OLE objects | 6 | `word/embeddings/oleObject1–6.bin` |

---

## Established APIs

### `DocxExtractor`

```python
from docx_extractor import DocxExtractor

result = DocxExtractor('template.docx').extract(
    output_dir='data',    # writes extraction.json here
    unzip_dir='template'  # reuses existing unzip if dir exists
)
# result is also returned as a dict
```

`result` keys: `source`, `extracted_at`, `metadata`, `placeholders`,
`sections`, `body_elements`, `headers`, `footers`, `relationships`

Each `body_elements` entry is one of:
- `{type: 'paragraph', index, style, text, runs, pPr, section_break}`
- `{type: 'table', index, rows: [[{text, paragraphs}]]}`
- `{type: <other_tag>, index}`

### `DocxRestorer`

```python
from docx_restore import DocxRestorer

DocxRestorer(
    template_dir='template',
    data={'field': 'value', ...}
).restore(output_path='output.docx')
```

`data` is a flat `str → str` dict. Keys must match placeholder names exactly
(case-sensitive). Placeholders in headers and footers are also substituted.

---

## Known Limitations (to address in advanced features)

| Limitation | Impact | Notes |
|---|---|---|
| No placeholders in template yet | Engine works but produces unmodified output | Template needs `{{field}}` annotations added |
| Substitution is paragraph-scoped | Cannot replace text that spans a paragraph boundary | By design — cross-paragraph replacement requires structural modification |
| Text inside drawing/shape text boxes not substituted | Shapes use `wps:txbx` subtree, not `w:r` | Needs separate shape-text traversal |
| No paragraph add/remove | Document length is fixed | Requires sectPr rebalancing if paragraphs shift sections |
| Table cell count is fixed | Cannot add/remove rows | Needs `w:tr` / `w:tc` cloning logic |
| `w:rPr` of replacement text is inherited from run[0] | Cannot inject bold, color, etc. via data | Needs a richer data model (e.g. `{text, bold, color}`) |
| No multi-document batch API | Must instantiate `DocxRestorer` per document | Easy to add as a thin loop wrapper |

---

## Logical Next Steps for Advanced Features

These are ordered roughly by dependency — each builds on the one before it.

### 1. Annotate the template with `{{placeholders}}`

Before any advanced feature is useful, the template text needs `{{field_name}}`
markers inserted at the right positions. The cover page table (body index 58)
and the title paragraphs are the primary targets.

Suggested fields based on the template structure:

```
{{title_zh}}         题目（中文）
{{title_en}}         题目（英文）
{{student_name}}     学生姓名
{{student_id}}       学号
{{supervisor}}       指导教师
{{major}}            专业
{{department}}       学院
{{year}}             年
{{month}}            月
{{day}}              日
{{abstract_zh}}      摘要正文（中文）
{{keywords_zh}}      关键词（中文）
{{abstract_en}}      Abstract body
{{keywords_en}}      Keywords (English)
```

### 2. Structured data model (`data/data.json` schema)

Define a canonical JSON schema for a thesis record so that data validation
can happen before restoration. Gives a stable contract between the data source
and the engine.

### 3. Multi-run formatting injection

Currently `_substitute_paragraph` collapses all runs into one. A richer
replacement could accept `{text, bold, italic, color}` and rebuild `w:rPr`
accordingly, enabling dynamic formatting of injected content.

### 4. Table row cloning

The three data tables (body indices 190, 206, 215) have fixed row counts.
A `clone_row(table_index, source_row, data_list)` utility would duplicate a
`w:tr` template row N times and fill each with data — needed for variable-
length result tables.

### 5. Shape / text-box substitution

OLE-embedded charts and WMF images with text labels currently cannot have
their text changed. For shapes that use `wps:txbx`, a separate traversal of
`w:drawing → wp:inline/wp:anchor → … → wps:txbx → w:txbxContent → w:p`
would extend substitution coverage to text boxes.

### 6. Header/footer dynamic content

Section headers show the chapter title (e.g., "第4章 …"). Currently these are
static in the XML. A `set_header_text(section_index, text)` method could
target the correct `headerN.xml` for a given section number using the section
map above.

### 7. Batch generation API

A thin wrapper around `DocxRestorer` to generate multiple documents from a
list of data dicts:

```python
def batch_restore(template_dir, records: list[dict], output_dir: str): ...
```

---

## Confidence in the Foundation

The extractor and restorer have been verified with:

- A live round-trip test injecting `{{student_name}}` into `document.xml`,
  running `DocxRestorer`, and confirming the replacement in the ZIP output.
- Confirmed that `template/` is never modified (original preserved after test).
- Confirmed `output.docx` is a valid ZIP with correct OOXML structure.
- All 58 relationships intact (headers, footers, images, OLE, styles, theme).

The lossless mechanism (namespace registration, XML declaration preservation,
run-merge strategy, temp-copy isolation) is documented in `DESIGN.md`.
