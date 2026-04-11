# CLAUDE.md

This file provides guidance to Claude Code when working with this repository.

---

# Project Overview

This repository contains **DOCX Engine V3**, a Python system for **programmatically extracting, representing, and restoring structured data from an academic Word template** used by Harbin Institute of Technology (哈尔滨工业大学).

The engine operates directly on **OOXML** rather than relying on high-level libraries.

The core idea is:

```
DOCX → XML → Structured Representation → Modified XML → DOCX
```

Two main workflows exist:

1. **Extraction**
   Parse a `.docx` template and extract its structural content.

2. **Restoration / Generation**
   Use the extracted representation to fill data and generate a new `.docx`.

---

# Repository Architecture

Core scripts:

```
docx_extractor.py
docx_restore.py
```

### docx_extractor.py

Responsible for converting DOCX content into structured data.

Main responsibilities:

* unzip DOCX
* parse XML files
* detect structural elements
* produce a structured representation

Typical extracted elements include:

* paragraphs
* runs (`w:r`)
* text nodes (`w:t`)
* headers / footers
* embedded objects
* images
* document styles

---

### docx_restore.py

Responsible for generating a new document using extracted structure.

Main responsibilities:

* load structured representation
* fill template placeholders
* reconstruct XML
* repack DOCX

The restore process **must preserve the original OOXML structure** to ensure compatibility with Word/WPS.

---

# DOCX Internal Structure

A `.docx` file is a **ZIP archive of OOXML files**.

Workflow used in this repository:

```
template.docx
    ↓ unzip
template/
    ↓ parse / modify XML
template/
    ↓ zip
output.docx
```

The `template/` directory mirrors the OOXML package structure.

Important files:

```
word/document.xml
```

Main body of the document.
Contains most paragraphs and content.

```
word/header*.xml
word/footer*.xml
```

Headers and footers for each section.

The template contains:

* 20 headers
* 16 footers

These are referenced through `rId` relations in:

```
word/_rels/document.xml.rels
```

Other components:

```
word/styles.xml
```

Defines all paragraph and character styles.

```
word/settings.xml
```

Document-level configuration.

```
word/media/*
```

Embedded images.

```
word/embeddings/*
```

OLE objects such as Excel charts or equation objects.

```
docProps/*
```

Document metadata.

---

# Template Characteristics

The template has several notable characteristics:

* Created with **WPS Office**
* Chinese locale (UTF-8 text)
* Multi-section document
* Heavy use of headers and footers
* Contains embedded OLE objects
* Uses mixed fonts (Times New Roman / 宋体)

The system must preserve these properties when generating new documents.

---

# Key OOXML Namespaces

Common namespaces used in the XML files:

```
w:  http://schemas.openxmlformats.org/wordprocessingml/2006/main
r:  http://schemas.openxmlformats.org/officeDocument/2006/relationships
v:  VML drawing namespace
w14: Word 2010 extensions
```

Most content manipulation occurs in `w:` elements.

---

# Development Commands

Run extractor:

```
python docx_extractor.py
```

Run restoration / generation:

```
python docx_restore.py
```

Manually unzip a docx:

```
python -c "import zipfile; zipfile.ZipFile('template.docx').extractall('template')"
```

Repack directory to DOCX:

```
python -c "
import zipfile, os
with zipfile.ZipFile('output.docx', 'w', zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk('template'):
        for f in files:
            p = os.path.join(root, f)
            z.write(p, os.path.relpath(p, 'template'))
"
```

---

# Data Directory

```
data/
```

Working directory for structured extraction results and generated data.

---

# Coding Guidelines

When modifying this repository:

1. **Do not break OOXML structure**

Never remove required nodes like:

* `w:body`
* `w:p`
* `w:r`
* relationship references (`rId`)

2. **Preserve namespaces**

XML namespaces must remain intact.

3. **Avoid formatting loss**

When modifying text content, preserve:

* run properties (`w:rPr`)
* paragraph properties (`w:pPr`)
* style references

4. **Prefer structural parsing**

Avoid fragile string replacements in XML.

Always parse XML nodes properly.

5. **Maintain deterministic output**

Generated documents should remain stable between runs.

---

# Typical Tasks Claude May Perform

Claude Code may help with:

* improving XML parsing logic
* adding new extraction fields
* fixing DOCX reconstruction bugs
* improving placeholder filling logic
* refactoring the extractor / restore pipeline

---

# Important Notes

The engine interacts with **raw OOXML**, not high-level document APIs.

Changes must be **structurally safe** to avoid corrupting DOCX files.

Always test changes by opening the generated document in Word or WPS.

---
