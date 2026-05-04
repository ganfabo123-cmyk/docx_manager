# DOCX JSON Structure

这个文档描述了 `docx_manager/docx_engine/engine` 中几阶段生成与消费的 JSON 结构，目的是让后续开发者通过文件格式而不是直接读代码来理解数据流。

## 1. 流程概览

主要涉及三个阶段：

1. `user_data_generator.py`
   - 输入：`data/full_parsed.json` + `sections_config/hit_config.json`
   - 输出：`data/user_data.json`
   - 作用：把内容数据和节配置信息转换成一系列可执行的 `docx_tools` 操作。

2. `user_data_compiler.py`
   - 输入：`data/user_data.json`
   - 读取：`data/extraction.json`（作为模板 scaffold）
   - 输出：`data/user_extraction.json`
   - 作用：执行 `docx_tools` 操作，生成完整的 extraction-style JSON。

3. `docx_compiler.py`
   - 输入：`data/user_extraction.json` 或其他 extraction-style JSON
   - 读取：`template/` 目录中的 DOCX 模板结构
   - 输出：`.docx`
   - 作用：把 extraction JSON 转换为真正的 Word 文档。

另外，`docx_tools.py` 提供了生成 `body_elements` 的构造函数；`save_document()` 输出的 JSON 只包含 `body_elements`，而不是完整 extraction scaffold。

---

## 2. `full_parsed.json` 规范

`user_data_generator.py` 读取的源数据允许两种格式：

- 规范格式：
  - `{"type": "abstract_cn", "data": {"content": "...", "keywords": [...]}}`
  - `{"type": "abstract_en", "data": {"content": "...", "keywords": [...]}}`
  - `{"type": "h1", "data": "章节标题"}`
  - `{"type": "h2", "data": "节标题"}`
  - `{"type": "h3", "data": "小节标题"}`
  - `{"type": "h4", "data": "小小节标题"}`
  - `{"type": "text", "data": "正文段落"}`
  - `{"type": "figure", "data": {"drawing_xml": "...", "caption": "..."}}`
  - `{"type": "table", "data": {"rows": [[...], ...], "caption": "..."}}`
  - `{"type": "equation", "data": {"expression": "...", "suffix": "(1-1)"}}`
  - `{"type": "reference", "data": {"index": 1, "content": "...", "before": "...", "after": "..."}}`

- Markdown 兼容格式：
  - `{"type": "body", "content": "# 一级标题"}` → `h1`
  - `{"type": "body", "content": "## 二级标题"}` → `h2`
  - `{"type": "body", "content": "### 三级标题"}` → `h3`
  - `{"type": "body", "content": "正文文本"}` → `text`

`user_data_generator.py` 内部会将这些输入正规化为形如 `{"type": <type>, "data": <str|dict>}` 的 canonical form。

---

## 3. `user_data.json` 结构（Action list）

`user_data_generator.py` 输出的文件结构为：

```json
{
  "document": [
    {"type": "add_heading", "text": "...", "level": 1},
    {"type": "add_paragraph", "text": "...", "style_type": "正文"},
    {"type": "generate_toc", "max_level": 4},
    {"type": "insert_page_break"},
    {"type": "insert_section_break", "header_refs": {...}, "footer_refs": {...}, "restart_page_number": 1},
    {"type": "insert_figure", "base64": "...", "width": 400, "height": 300, "caption": "...", "position": "center"},
    {"type": "insert_table", "rows": [[...], ...], "caption": "...", "auto_format": true, "column_widths": [...]},
    {"type": "insert_equation", "expression": "...", "category": "omath", "omml": "...", "suffix": "...", "label": "...", "is_inline": false},
    {"type": "add_reference", "index": 1, "content": "...", "before": "...", "after": "...", "auto_cite": false},
    {"type": "insert_abstract_with_keywords", "cn_content": "...", "en_content": "...", "cn_keywords": [...], "en_keywords": [...], "cn_section_break": {...}, "en_section_break": {...}}
  ]
}
```

### 支持的 action 类型

- `add_heading`：参数 `text`, `level`
- `add_paragraph`：参数 `text`, `style_type`
- `generate_toc`：参数 `max_level`
- `insert_page_break`：无参数
- `insert_section_break`：可选参数 `header_template`, `footer_template`, `restart_page_number`, `header_refs`, `footer_refs`
- `insert_figure`：参数包括 `data_source` / `base64` / `drawing_xml`, `width`, `height`, `caption`, `position`
- `insert_table`：参数包括 `rows`, `caption`, `auto_format`, `column_widths`
- `insert_equation`：参数包括 `expression`, `category`, `omml`, `ole_base64`, `image_base64`, `label`, `width_pt`, `height_pt`, `prog_id`, `text_before`, `text_after`, `is_inline`
- `insert_abstract_with_keywords`：参数包括中文/英文摘要文本、关键词、标题、关键字标签、以及对应的 `cn_section_break` / `en_section_break`
- `add_reference`：参数包括 `index`, `content`, `before`, `after`, `auto_cite`

### 重点说明

- `user_data_generator.py` 会根据 `full_parsed.json` 中的 `figure`、`table`、`equation`、`reference` 等类型生成对应 action。
- `add_reference` 的 `before` / `after` 只是元数据，当前仅用于后处理或引用上下文，不影响 Word 的基本排版。

---

## 4. `user_extraction.json` / extraction-style JSON 结构

`user_data_compiler.py` 生成的 `data/user_extraction.json` 采用与 `extraction.json` 相同的 top-level 结构：

```json
{
  "source": "template.docx",
  "headers": {...},
  "footers": {...},
  "relationships": {...},
  "sections": [...],
  "body_elements": [...]
}
```

### 顶层字段说明

- `source`：固定字符串 `template.docx`
- `headers`：从原始 `extraction.json` 直接复制的 header 数据，保持不变
- `footers`：从原始 `extraction.json` 直接复制的 footer 数据，保持不变
- `relationships`：从原始 `extraction.json` 直接复制，用于编译器解析 rId
- `sections`：由 `body_elements` 中含 `section_break` 的段落派生出来
- `body_elements`：由 `docx_tools` 生成的实际文档内容元素列表

### `sections` 条目结构

每个 section 条目来源于一个段落的 `section_break`：

```json
{
  "paragraph_index": 123,
  "header_refs": {"default": "rId10", "even": "rId11"},
  "footer_refs": {"default": "rId12"},
  "page_size": {"w": "11906", "h": "16838"},
  "restart_page_number": 1
}
```

- `paragraph_index`：对应 `body_elements` 中发生换节的段落索引
- `header_refs` / `footer_refs`：rId 映射，或 `__template__` 占位符
- `page_size`：页面尺寸，字符串形式（默认 A4）
- `restart_page_number`：可选，表示节内页码重置

### `headers` / `footers` / `relationships`

这三组字段在当前代码里不是由 `user_data_compiler.py` 生成的，而是直接从 `data/extraction.json` 复制过来。
它们对 compiler 来说是“黑盒”数据：只要原始 scaffold 结构与 `docx_compiler.py` 期望一致，编译器即可正常工作。

---

## 5. `body_elements` 详细结构

`body_elements` 是整个 pipeline 的核心数据结构。当前支持的 element types 有：

- `paragraph`
- `toc`
- `raw_xml`
- `table`
- `image`
- `omath` / `omathpara`
- `ole`

### paragraph

```json
{
  "index": 0,
  "type": "paragraph",
  "style": null,
  "text": "正文内容...",
  "runs": [
    {"text": "第一段", "rPr": "<ns0:rPr ...>...</ns0:rPr>"},
    {"text": "[1]", "rPr": "<ns0:rPr ...>...</ns0:rPr>"}
  ],
  "pPr": "<ns0:pPr ...>...</ns0:pPr>",
  "section_break": {
    "header_refs": {"default": "__template__"},
    "footer_refs": {"default": "__template__"},
    "page_size": {"w": "11906", "h": "16838"},
    "restart_page_number": 1
  }
}
```

- `index`：按生成顺序分配的唯一段落编号
- `type`：固定为 `paragraph`
- `style`：段落样式 ID（例如 `"2"`、`"3"`），也可以为 `null`
- `text`：段落文本的扁平化表示
- `runs`：文本段的列表，保留每个 run 的 `text` 和 `rPr`
- `pPr`：段落级 OOXML XML 字符串
- `section_break`：可选，描述该段落后的节信息

#### run 对象

```json
{
  "text": "...",
  "rPr": "<ns0:rPr ...>...</ns0:rPr>",
  "break_type": "page",
  "drawing_xml": "...",
  "object_xml": "..."
}
```

- `text`：运行文本
- `rPr`：运行样式 XML 字符串
- `break_type`：可选，仅用于 `insert_page_break`
- `drawing_xml`：可选，用于原始图像 DrawingML
- `object_xml`：可选，用于原始 OLE 对象 XML

### toc

```json
{
  "index": 12,
  "type": "toc",
  "max_level": 4
}
```

- `max_level`：目录最大包含的标题级别
- 由 `docx_tools.generate_toc()` 生成，`docx_compiler.py` 会把它转为 TOC 域

### raw_xml

```json
{
  "index": 13,
  "type": "raw_xml",
  "xml": "<w:sdt>...</w:sdt>"
}
```

- 这个类型用于直接重放 extractor 保存的原始 XML 片段，编译器不对其语义做解析。

### table

```json
{
  "index": 14,
  "type": "table",
  "rows": [
    [
      {
        "text": "单元格文本",
        "paragraphs": [ ... ]
      }
    ]
  ],
  "auto_format": true,
  "column_widths": [50.0, 50.0]
}
```

- `rows`：二维列表，每个单元格是一个 dict
- `auto_format`：是否启用自动格式化
- `column_widths`：可选列宽比例/绝对值

#### table cell

```json
{
  "text": "...",
  "paragraphs": [
    {
      "index": 0,
      "type": "paragraph",
      "style": null,
      "text": "...",
      "runs": [...],
      "pPr": "...",
      "section_break": null
    }
  ]
}
```

- 每个单元格保持至少一个段落
- `paragraphs` 内的段落结构与 body element 的 paragraph 相同

### image

两类 image 元素：

- `drawing_xml` 路径：保留原始 DrawingML，适用于从 extractor 继承的图片
- `base64` 路径：由 `docx_tools.insert_figure()` 新增的图片

```json
{
  "index": 15,
  "type": "image",
  "drawing_xml": "<w:drawing>...</w:drawing>",
  "caption": "图注",
  "position": "center"
}
```

或

```json
{
  "index": 16,
  "type": "image",
  "base64": "...",
  "width": 400,
  "height": 300,
  "caption": "图注",
  "position": "center"
}
```

- `caption`：图片说明文本，`docx_tools` 通常会额外生成一个 caption 段落
- `position`：`left` / `center` / `right`

### omath / omathpara

```json
{
  "index": 17,
  "type": "omath",
  "formula": "<m:oMath>...</m:oMath>",
  "formula_index": "(1-1)",
  "position": "center",
  "suffix_position": "right",
  "is_inline": true,
  "text_before": "式中",
  "text_after": "..."
}
```

- `formula`：OMML XML 或表达式文本
- `formula_index`：公式编号
- `position` / `suffix_position`：布局控制
- `is_inline`、`text_before`、`text_after`：控制行内公式行为

### ole

```json
{
  "index": 18,
  "type": "ole",
  "formula": "...",
  "formula_index": "(1-1)",
  "position": "center",
  "suffix_position": "right",
  "base64": "...",
  "image_base64": "...",
  "width_pt": 120,
  "height_pt": 30,
  "prog_id": "Equation.3"
}
```

- `base64`：OLE 对象二进制数据
- `image_base64`：OLE 预览图（可选）
- `prog_id`：OLE 对象 ProgID

---

## 6. `docx_tools.py` 生成规则

`docx_tools.py` 的每个公开函数最终都会把一个 body element 追加到模块级 `_DOCUMENT` 列表。

- `save_document(path)`：把 `_DOCUMENT` 序列化为 `{"body_elements": _DOCUMENT}`
- `get_document()`：取当前 body elements
- `clear_document()`：重置文档状态

`docx_tools` 与 extraction JSON 的对应关系：

- `add_heading()` → `paragraph` + style ID
- `add_paragraph()` → `paragraph` + `style=null` + `pPr=_BODY_PPR`
- `generate_toc()` → `toc`
- `insert_page_break()` → `paragraph` + run `break_type="page"`
- `insert_section_break()` → `paragraph` + `section_break` dict
- `insert_figure()` → `image` (+ optional caption paragraph)
- `insert_table()` → `table` (+ optional caption paragraph)
- `insert_equation()` → `omath`/`ole`
- `insert_abstract_with_keywords()` → 一组 `paragraph` 和关键词段落，部分段落携带 `section_break`
- `add_reference()` → `paragraph` + 额外字段 `citation_index` / `citation_before` / `citation_after` / `auto_cite`

---

## 7. `docx_compiler.py` 消费规则

`docx_compiler.py` 读取 extraction-style JSON 后，按顺序重建 `word/document.xml`：

- `paragraph` → `_build_para()`
- `raw_xml` → 直接 `_parse_xml()` 并插入
- `toc` → `_build_toc_para()`
- `table` → `_build_table()`
- `image` → `_build_image_nodes()`
- `omath` / `omathpara` → `_build_omath_para()`
- `ole` → `_build_ole_para()`

### 重点行为

- `section_break` 会被 `_build_sectPr()` 插入到对应段落的 `pPr` 中。
- `drawing_xml` 原样保留，必要时会自动补充 `mc:AlternateContent`。 
- `raw_xml` 用于保留 TOC、结构化域、Sdt 等复杂原始 XML。
- `table` 的结构骨架来自模板；如果模板中没有足够的表格，编译器会生成最小合法表结构。
- `body-level sectPr` 来自模板而非 extraction JSON，因此不是 `section_break` 的全部信息。

---

## 8. 关键约定

- `index` 字段仅用于内部排序/调试，编译器不会把它写入 Word 内容。
- `text` 字段是对段落内容的扁平化表示，实际排版以 `runs` 中的 `text` 和 `rPr` 为准。
- `style` 字段可以为 `null`，表示不依赖命名段落样式。
- `section_break` 只影响后续段落起始的新节；如果最后一个 element 是 `section_break`，`user_data_generator.py` 会尝试删除它以避免空白页。
- `headers`/`footers`/`relationships` 在当前 pipeline 中只做原样透传，不参与生成。

---

## 9. 参考关系

- `user_data_generator.py` → 生成 `data/user_data.json`
- `user_data_compiler.py` → 将 `data/user_data.json` 编译为 `data/user_extraction.json`
- `docx_tools.py` → 提供 `body_elements` 构造函数与 `save_document()`
- `docx_compiler.py` → 消费 `data/user_extraction.json` / extraction JSON 并输出 `.docx`
