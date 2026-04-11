

# docx_tools Implementation Specification

## Stage Goal

The goal of this stage is to **implement Python tools that generate `extraction.json` entries** for a DOCX document. Each function corresponds to a specific type of content element, preserving the linear structure and style extracted by the extractor.

All functions operate **linearly**, from top to bottom, mirroring the order in the original document. Each function only requires the **core user content** as input and generates a `body_element` entry for `xxxx_extraction.json`.

---

## General Principles

1. **Invariant Principle:** Each function accepts **user content only**. Formatting, styling, and layout are automatically applied from the extractor’s data.
2. **Linear Mapping:** The function call sequence should reflect the linear order of `body_elements`:

   ```
   HEADING1(...)
   HEADING2(...)
   TOC(...)
   SECTIONPTR(...)
   NORMAL(...)
   IMAGE(...)
   TABLE(...)
   OLE(...)
   ```
3. **No manual formatting or XML manipulation:** All functions generate JSON entries consistent with `extraction.json` format.

---

## Function Specifications

### 1. Headings (`heading1`, `heading2`, …, `headingN`)

* **Input:**

  * `text`: heading text
  * `level`: optional, heading level (if using a unified function)
* **Output:** JSON `body_element` representing the heading
* **Behavior:**

  * Automatically applies the correct style and index from the extractor
  * Preserves section hierarchy for TOC and section references

---

### 2. Table of Contents (`toc` / `table_of_contents`)

* **Input:**

  * `titles_list`: list of all document headings in order
  * `toc_title`: e.g., `"Table of Contents"` or `"目 录"`
* **Output:** JSON `body_element` representing a TOC
* **Behavior:**

  * Generates an **auto-updating field** (Word’s built-in TOC field)
  * No hyperlinks or manual page numbers are needed; Word updates automatically
  * Should be placed after the first heading, before the first section content

---

### 3. Section Pointer (`section_ptr`)

* **Input:**

  * Reference to a heading or section
* **Output:** JSON `body_element` marking the section location
* **Behavior:**

  * Serves as a logical anchor for cross-references or navigation
  * Optional, only if needed for internal references

---

### 4. Normal Text (`normal_text`)

* **Input:**

  * `text`: paragraph content (may contain line breaks)
* **Output:** JSON `body_element` representing a paragraph
* **Behavior:**

  * Preserves extractor paragraph styles
  * No additional formatting required

---

### 5. Table (`table`)

* **Input:**

  * `rows`: 2D array representing table content
  * `heading`: optional table caption/title
* **Output:** JSON `body_element` for a table
* **Behavior:**

  * Maintains original style and layout from the extractor
  * Each row/column maps to JSON entries as per extraction format

---

### 6. Image (`image`)

* **Input:**

  * `base64_str`: base64-encoded image content
  * `heading`: optional image caption
  * `position`: `"center"`, `"left"`, `"right"`
  * `width` / `height`: pixel or EMU units
* **Output:** JSON `body_element` representing the image
* **Behavior:**

  * Supports inline and anchor images automatically based on content
  * Image metadata (position, size, caption) is preserved

---

### 7. OLE / OMath / OMathPara (`ole`, `omath`, `omathpara`)

* **Input:**

  * `formula_str`: formula content (OLE or OMath XML)
  * `index`: formula index, e.g., `(4-2)`
* **Output:** JSON `body_element` for the formula
* **Behavior:**

  * Inline or floating formula automatically preserved
  * Index ensures correct order and reference in extraction JSON

---

## Example Linear Call Sequence

```python
HEADING1("Chapter 1: Overview")
TOC(["Chapter 1: Overview", "Chapter 2: Method"])
SECTIONPTR("Chapter 1")
NORMAL("This is the introductory paragraph...")
IMAGE(base64_str, "Figure 1", "center", 400, 300)
TABLE([["Header1","Header2"], ["Data1","Data2"]], "Table 1")
OLE("formula content", "(1-1)")
```

* This sequence mirrors the order of the original document
* Each function generates a single `body_element` for `xxxx_extraction.json`
* TOC is automatically generated from headings, no manual XML or hyperlink management

