# Design Notes: Lossless DOCX Parsing and Reconstruction

This document explains the key design decisions behind `docx_extractor.py` and
`docx_restore.py`, with a focus on how the engine achieves lossless round-trips
through a `.docx` file without corrupting formatting, structure, or binary assets.

---

## 1. What "Lossless" Means Here

A lossless round-trip means:

```
template.docx  →  extract  →  restore  →  output.docx
```

where `output.docx` opens in Word or WPS with:

- **identical visual layout** (fonts, sizes, spacing, indentation)
- **identical document structure** (sections, headers, footers, page sizes)
- **intact binary assets** (OLE objects, images, embedded charts)
- **intact relationships** (every `rId` still points to the correct target)

The only permitted difference is the substituted text content.

The main risk is accidental loss. Every design decision below exists to prevent
a specific class of loss.

---

## 2. The DOCX Container: Don't Touch What You Don't Need To

A `.docx` file is a ZIP archive. Its contents look like:

```
[Content_Types].xml
_rels/.rels
word/document.xml          ← main body
word/styles.xml
word/settings.xml
word/header1.xml … header20.xml
word/footer1.xml … footer16.xml
word/theme/theme1.xml
word/media/image1.tiff … image5.tiff
word/media/image6.wmf  … image11.wmf
word/embeddings/oleObject1.bin … oleObject6.bin
word/_rels/document.xml.rels
docProps/app.xml
docProps/core.xml
```

**The first lossless principle: only rewrite files you actually modified.**

The restorer copies the entire `template/` directory to a temporary working
copy, then calls `_process_xml_file` only on `document.xml` and the
header/footer XML files. Every other file — styles, settings, theme, font
table, all binary media, all OLE embeddings, `[Content_Types].xml`,
relationship files — is copied verbatim and repacked unchanged.

This means corruption of images, embedded Excel charts, or equation objects is
structurally impossible: those files are never opened or touched.

---

## 3. The XML Layer: Surgical Edits via ElementTree

### 3.1 Parse → modify in-place → serialize

The restorer uses Python's `xml.etree.ElementTree` to parse each XML file
into a live element tree, mutates only the specific text nodes that need
changing, and serializes the whole tree back to disk.

```python
tree = ET.parse(xml_path)
root = tree.getroot()
# … modify only w:t text nodes …
xml_body = ET.tostring(root, encoding='unicode')
```

Compared to string/regex replacement on raw XML, this approach:

- Cannot accidentally corrupt a tag name, attribute value, or namespace declaration
- Cannot introduce an unbalanced `<` or `>` that breaks XML well-formedness
- Cannot touch any element other than the ones explicitly targeted

### 3.2 Namespace prefix preservation

OOXML files declare around 17 namespaces on the root element:

```xml
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  … 13 more …
>
```

Python's ElementTree stores tags and attributes in **Clark notation** — the
URI replaces the prefix:

```
w:document  →  {http://schemas.openxmlformats.org/wordprocessingml/2006/main}document
w14:paraId  →  {http://schemas.microsoft.com/office/word/2010/wordml}paraId
```

When `ET.tostring()` serializes the tree, it looks up the registered prefix
for each URI. If a URI has no registered prefix, ElementTree invents one:
`ns0:`, `ns1:`, etc. Word and WPS would fail to open such a file because they
match namespace URIs by prefix name in some legacy code paths, and because
`mc:Ignorable="w14 w15 wp14"` references prefixes by name.

The fix is to call `ET.register_namespace()` for every URI **before** parsing:

```python
_NS_MAP = {
    'w':    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'w14':  'http://schemas.microsoft.com/office/word/2010/wordml',
    'mc':   'http://schemas.openxmlformats.org/markup-compatibility/2006',
    # … all 17 …
}
for prefix, uri in _NS_MAP.items():
    ET.register_namespace(prefix, uri)
```

After this, `ET.tostring()` uses `w:`, `r:`, `w14:`, etc. — exactly matching
the original. The `xml:` prefix (`http://www.w3.org/XML/1998/namespace`) is
built into the XML specification and into ElementTree; it does not need
registration and is always serialized correctly.

### 3.3 Preserving the XML declaration

ElementTree does not store the `<?xml … ?>` processing instruction as part of
the element tree. If you use `ET.write(file, xml_declaration=True)`, the
output is:

```xml
<?xml version='1.0' encoding='us-ascii'?>
```

That uses single quotes and omits `standalone="yes"`, both of which differ
from the original. OOXML convention (and WPS's parser) expects:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
```

The solution is to prepend the declaration manually as a string:

```python
with open(xml_path, 'w', encoding='utf-8') as f:
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    f.write(ET.tostring(root, encoding='unicode'))
```

`encoding='unicode'` tells `ET.tostring` to return a Python `str` (not bytes),
avoiding any encoding mismatch. The file is then written as UTF-8 bytes,
consistent with the declaration.

---

## 4. The Run-Merge Problem

### 4.1 Why Word splits text across runs

In OOXML, a paragraph (`w:p`) contains one or more runs (`w:r`), each of which
holds a text node (`w:t`) and optional run properties (`w:rPr`). A run
boundary exists wherever formatting changes.

Word and WPS add run boundaries for reasons beyond visible formatting — spell
checker state, revision marks, autocorrect markers, and even random internal
bookkeeping. The result is that a string you type as one unit may be stored as
several consecutive runs with identical formatting:

```xml
<w:p>
  <w:r><w:t>{{stu</w:t></w:r>
  <w:r><w:t>dent</w:t></w:r>
  <w:r><w:t>_name}}</w:t></w:r>
</w:p>
```

A naive `str.replace('{{student_name}}', value)` on either the raw XML or on
per-run text would miss this completely.

### 4.2 The merge-then-rebuild strategy

`_substitute_paragraph` handles this in three steps:

**Step 1 — collect run slots:**

```python
run_slots = []
for r in runs:
    t = r.find(_q('t'))
    run_slots.append((r, t, t.text or '' if t is not None else ''))
```

Each slot records the run element, its `w:t` child, and the text as a plain
string.

**Step 2 — merge and substitute on the full text:**

```python
full_text = ''.join(text for _, _, text in run_slots)
new_text  = full_text
for key, value in data.items():
    new_text = new_text.replace(f'{{{{{key}}}}}', str(value))
```

The placeholder is matched against the merged string, so split-run placeholders
are found regardless of how many fragments they were split into.

**Step 3 — rebuild into first run, hollow out the rest:**

```python
first_t.text = new_text         # full replacement in run[0]

for _, t, _ in run_slots[1:]:
    if t is not None:
        t.text = ''             # empty, not removed
```

The remaining runs are emptied but **not deleted**. Their `w:r` elements and
`w:rPr` children remain in the tree. This is critical: removing a run that
carries a spell-check or revision mark could shift internal bookmarks or
corrupt tracked changes. An empty run is inert but structurally safe.

The replaced text inherits the `w:rPr` of the first run. For a placeholder,
this is correct: the placeholder was written with the intended formatting of
the final output.

### 4.3 Whitespace preservation

XML parsers strip leading and trailing whitespace from text nodes by default.
OOXML signals "do not strip" via `xml:space="preserve"` on the `w:t` element.
After substitution, the attribute is set or cleared based on the actual content:

```python
if new_text and (new_text[0] == ' ' or new_text[-1] == ' '):
    first_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
else:
    first_t.attrib.pop('{http://www.w3.org/XML/1998/namespace}space', None)
```

Forgetting this causes Word to silently drop leading/trailing spaces on load —
a subtle loss that would be hard to diagnose.

---

## 5. Structure Never Reconstructed from Scratch

A key architectural decision is that the restorer **never reconstructs XML from
scratch**. It always starts from the original template tree and mutates only
text nodes.

The alternative — building new `w:p` / `w:r` / `w:t` trees programmatically —
would require replicating every `w:pPr`, `w:rPr`, `w:sectPr`, spacing
attribute, indentation value, border definition, and so on. Omitting even one
non-obvious attribute (e.g., `w:adjustRightInd`, `w:snapToGrid`,
`w:contextualSpacing`) silently changes paragraph layout. The template already
encodes all of this correctly; the engine inherits it for free by preserving
the tree.

This is why the extraction JSON stores `pPr` and `rPr` as serialized XML
strings rather than decoded dicts: they are meant to be inspectable, not
re-ingested for reconstruction.

---

## 6. Section and Header/Footer Integrity

The document has 11 sections, each referencing up to 3 headers (first, default,
even) and up to 2 footers by `rId`. These relationships live in
`word/_rels/document.xml.rels` and are never modified by the restorer.
The header and footer XML files themselves are only modified if they contain
`{{placeholder}}` text — otherwise they pass through untouched as binary blobs
in the ZIP.

Section properties (`w:sectPr`) contain page size (`w:pgSz`), margin
definitions (`w:pgMar`), columns, line numbering, and header/footer distance.
They live either inside `w:pPr` of a section-break paragraph or directly as a
child of `w:body`. The run-merge function operates on `w:r` children of
`w:p`; it never descends into `w:pPr`, so `w:sectPr` is untouched by
substitution even when it lives inside a paragraph.

---

## 7. The Immutable Template Guarantee

The restorer always operates on a `tempfile.mkdtemp()` copy:

```python
tmp_base = tempfile.mkdtemp(prefix='docx_restore_')
work_dir = os.path.join(tmp_base, 'work')
shutil.copytree(self.template_dir, work_dir)
# … all modifications happen in work_dir …
_pack_docx(work_dir, output_path)
shutil.rmtree(tmp_base, ignore_errors=True)
```

`template/` is never touched. This means:

- The restorer is safe to call in a loop (generating many documents from one template).
- A crash or exception during processing leaves the template intact.
- The `finally` block ensures the temp directory is cleaned up even on error.

---

## 8. Deterministic ZIP Output

Python's `zipfile` does not guarantee a stable file order across runs. The
restorer enforces determinism by sorting both directory names and filenames
before adding them to the archive:

```python
for dirpath, dirnames, filenames in os.walk(source_dir):
    dirnames.sort()
    for filename in sorted(filenames):
        …
```

This means two calls with identical input produce bit-for-bit identical output
ZIP archives. Determinism matters for version control diffs and for verifying
that no unintended changes were introduced.

---

## 9. What the Engine Does Not Handle (Known Limits)

| Scenario | Status |
|---|---|
| Placeholder split across paragraphs | Not supported — substitution is paragraph-scoped |
| Replacing text inside a table cell | Supported — `root.iter(w:p)` recurses into tables |
| Replacing text inside a header/footer | Supported — headers/footers processed separately |
| Replacing text inside a drawing/shape text box | Not supported — shape text is in a different XML subtree |
| Modifying `w:rPr` formatting as part of substitution | Not supported — value is text-only |
| Adding or removing paragraphs | Not supported — tree structure is fixed |
| Binary assets (images, OLE objects) | Copied verbatim, never modified |
