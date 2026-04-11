"""
Extracts structured content from a .docx file and saves it as JSON.

Usage:
    python docx_extractor.py [input.docx] [output_dir]

Defaults:
    input      = template.docx
    output_dir = data/          (writes extraction.json)
"""

import json
import os
import re
import sys
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

# ── Namespace URIs ────────────────────────────────────────────────────────────
W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
MC  = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

# Field types that must be preserved as raw XML rather than parsed.
# Parsing w:fldChar-based structures (TOC, cross-refs, page numbers) is
# guaranteed to break them — preserve the complete XML block instead.
_PRESERVED_FIELDS = frozenset({'TOC', 'REF', 'SEQ', 'PAGE', 'NUMPAGES', 'STYLEREF'})


def _q(local: str) -> str:
    """Return Clark-notation tag for the w: namespace."""
    return f'{{{W}}}{local}'


def _tag(elem: ET.Element) -> str:
    """Return the local (unprefixed) tag name of an element."""
    t = elem.tag
    return t.split('}', 1)[1] if '}' in t else t


def _text_of(elem: ET.Element) -> str:
    """Concatenate all w:t text nodes under elem."""
    return ''.join(t.text or '' for t in elem.iter(_q('t')))


# ── Field block detection ──────────────────────────────────────────────────────

def _has_field_begin(p_elem: ET.Element) -> bool:
    """Return True if the paragraph contains a w:fldChar with fldCharType='begin'."""
    for fc in p_elem.iter(_q('fldChar')):
        if fc.get(_q('fldCharType')) == 'begin':
            return True
    return False


def _field_depth_delta(elem: ET.Element) -> int:
    """Net change in field nesting depth contributed by this element."""
    delta = 0
    for fc in elem.iter(_q('fldChar')):
        t = fc.get(_q('fldCharType'))
        if t == 'begin':
            delta += 1
        elif t == 'end':
            delta -= 1
    return delta


def _collect_field_block(
    children: list,
    start: int,
) -> 'tuple[list | None, int]':
    """
    Starting at children[start] (a paragraph with a fldChar begin), collect
    all elements up to and including the element that closes the outermost field.

    Returns (block, consumed) when the field instruction is in _PRESERVED_FIELDS.
    Returns (None, 0) otherwise so the caller falls through to normal extraction.
    """
    depth = 0
    block: list = []
    j = start

    while j < len(children):
        elem = children[j]
        block.append(elem)
        depth += _field_depth_delta(elem)
        j += 1
        if depth <= 0:
            break

    # Gather full instrText from all collected elements
    full_instr = ''.join(
        t.text or ''
        for elem in block
        for t in elem.iter(_q('instrText'))
    )
    first_word = full_instr.strip().split()[0].upper() if full_instr.strip() else ''

    if first_word in _PRESERVED_FIELDS:
        return block, j - start
    return None, 0


# ── Run extraction ─────────────────────────────────────────────────────────────

def _is_inline_drawing(drawing_xml: str) -> bool:
    """Return True when drawing_xml wraps a wp:inline element.

    wp:inline drawings are content images (raster pictures, etc.).
    wp:anchor drawings are floating shapes (lines, boxes, etc.) and should
    stay inside their paragraph rather than being promoted to image elements.
    """
    try:
        elem = ET.fromstring(drawing_xml)
        return elem.find(f'{{{WP}}}inline') is not None
    except ET.ParseError:
        return False


def _extract_run(r_elem: ET.Element) -> dict:
    rPr     = r_elem.find(_q('rPr'))
    t       = r_elem.find(_q('t'))
    drawing = r_elem.find(_q('drawing'))
    obj     = r_elem.find(_q('object'))

    base: dict = {
        'text': '',
        'rPr':  ET.tostring(rPr, encoding='unicode') if rPr is not None else None,
    }

    # WPS Office wraps anchor shapes (text boxes, connectors, lines) inside
    # mc:AlternateContent so that Word can fall back to a VML rendering.
    # We must store the ENTIRE mc:AlternateContent element — including the
    # mc:Fallback / VML section — because WPS needs the wrapper to honour the
    # anchor's posOffset values.  Storing only the inner w:drawing (the
    # previous behaviour) causes WPS to ignore the position and render the
    # shape at an arbitrary location.
    if drawing is None:
        alt = r_elem.find(f'{{{MC}}}AlternateContent')
        if alt is not None:
            choice = alt.find(f'{{{MC}}}Choice')
            if choice is not None and choice.find(_q('drawing')) is not None:
                base['drawing_xml'] = ET.tostring(alt, encoding='unicode')
                return base

    if drawing is not None:
        base['drawing_xml'] = ET.tostring(drawing, encoding='unicode')
        return base

    if obj is not None:
        # OLE embedded object (e.g. Equation Editor formula) — preserve the
        # full VML + OLEObject XML so the compiler can re-emit it verbatim.
        # The rId references (OLE binary + preview image) remain valid because
        # the compiler copies the template's embeddings/ and _rels/ unchanged.
        base['object_xml'] = ET.tostring(obj, encoding='unicode')
        return base

    base['text'] = t.text or '' if t is not None else ''
    return base


# ── Paragraph extraction ───────────────────────────────────────────────────────

def _extract_sectPr(sectPr: ET.Element) -> dict:
    header_refs, footer_refs = {}, {}
    for ref in sectPr.findall(_q('headerReference')):
        header_refs[ref.get(_q('type'), 'default')] = ref.get(f'{{{R}}}id')
    for ref in sectPr.findall(_q('footerReference')):
        footer_refs[ref.get(_q('type'), 'default')] = ref.get(f'{{{R}}}id')
    pgSz = sectPr.find(_q('pgSz'))
    return {
        'header_refs': header_refs,
        'footer_refs': footer_refs,
        'page_size':   {'w': pgSz.get(_q('w')), 'h': pgSz.get(_q('h'))} if pgSz is not None else None,
    }


def _extract_paragraph(p_elem: ET.Element, index: int) -> dict:
    pPr      = p_elem.find(_q('pPr'))
    style_id = None
    sect_info = None

    if pPr is not None:
        style_ref = pPr.find(_q('pStyle'))
        if style_ref is not None:
            style_id = style_ref.get(_q('val'))
        sectPr = pPr.find(_q('sectPr'))
        if sectPr is not None:
            sect_info = _extract_sectPr(sectPr)

    runs = [_extract_run(c) for c in p_elem if _tag(c) == 'r']

    return {
        'index':         index,
        'type':          'paragraph',
        'style':         style_id,
        'text':          _text_of(p_elem),
        'runs':          runs,
        'pPr':           ET.tostring(pPr, encoding='unicode') if pPr is not None else None,
        'section_break': sect_info,
    }


# ── Table extraction ───────────────────────────────────────────────────────────

def _extract_cell(tc: ET.Element) -> dict:
    paras = [_extract_paragraph(p, i) for i, p in enumerate(tc.findall(_q('p')))]
    return {'text': _text_of(tc), 'paragraphs': paras}


def _extract_table(tbl: ET.Element, index: int) -> dict:
    rows = [
        [_extract_cell(tc) for tc in tr.findall(_q('tc'))]
        for tr in tbl.findall(_q('tr'))
    ]
    return {'index': index, 'type': 'table', 'rows': rows}


# ── Header / footer extraction ─────────────────────────────────────────────────

def _extract_hf_file(xml_path: str) -> dict | None:
    if not os.path.exists(xml_path):
        return None
    root = ET.parse(xml_path).getroot()
    paras = [_extract_paragraph(p, i) for i, p in enumerate(root.iter(_q('p')))]
    return {'text': _text_of(root), 'paragraphs': paras}


# ── Relationships ──────────────────────────────────────────────────────────────

def _load_relationships(rels_path: str) -> dict:
    if not os.path.exists(rels_path):
        return {}
    root = ET.parse(rels_path).getroot()
    result = {}
    for rel in root.findall(f'{{{REL}}}Relationship'):
        rtype = rel.get('Type', '').rsplit('/', 1)[-1]
        result[rel.get('Id')] = {'type': rtype, 'target': rel.get('Target')}
    return result


# ── Metadata ───────────────────────────────────────────────────────────────────

def _load_metadata(docx_dir: str) -> dict:
    meta = {}
    for rel_path in ('docProps/app.xml', 'docProps/core.xml'):
        path = os.path.join(docx_dir, rel_path)
        if os.path.exists(path):
            for child in ET.parse(path).getroot():
                if child.text:
                    meta[_tag(child)] = child.text
    return meta


# ── Main extractor class ───────────────────────────────────────────────────────

class DocxExtractor:
    def __init__(self, docx_path: str = 'template.docx'):
        self.docx_path = docx_path

    def extract(self, output_dir: str = 'data', unzip_dir: str | None = None) -> dict:
        """
        Unzip the .docx (if needed), parse all content, save extraction.json.

        Args:
            output_dir: directory to write extraction.json
            unzip_dir:  where to unzip the docx (defaults to stem of docx filename)

        Returns:
            The extracted structure as a dict.
        """
        # 1. Unzip
        if unzip_dir is None:
            unzip_dir = os.path.splitext(os.path.basename(self.docx_path))[0]

        if not os.path.isdir(unzip_dir):
            print(f'[extractor] Unzipping {self.docx_path!r} → {unzip_dir}/')
            with zipfile.ZipFile(self.docx_path, 'r') as zf:
                zf.extractall(unzip_dir)
        else:
            print(f'[extractor] Using existing directory: {unzip_dir}/')

        word_dir  = os.path.join(unzip_dir, 'word')
        rels_path = os.path.join(word_dir, '_rels', 'document.xml.rels')

        # 2. Relationships
        relationships = _load_relationships(rels_path)

        # 3. Parse document.xml body
        body = ET.parse(os.path.join(word_dir, 'document.xml')).getroot().find(_q('body'))

        body_elements: list = []
        sections:      list = []

        children = list(body)
        i = 0
        while i < len(children):
            child = children[i]
            tag   = _tag(child)
            idx   = len(body_elements)

            # ── Field block preservation ───────────────────────────────────────
            # Paragraphs that start a preserved field (TOC, REF, PAGE, …) are
            # collected in their entirety and stored as raw_xml so the compiler
            # can emit them verbatim — parsing and regenerating fldChar structures
            # is guaranteed to break TOC hyperlinks, page numbers, etc.
            if tag == 'p' and _has_field_begin(child):
                block, consumed = _collect_field_block(children, i)
                if block is not None:
                    for blk_elem in block:
                        blk_idx = len(body_elements)
                        body_elements.append({
                            'index': blk_idx,
                            'type':  'raw_xml',
                            'xml':   ET.tostring(blk_elem, encoding='unicode'),
                        })
                    i += consumed
                    continue
                # Field type not in preserved list — fall through to normal extraction

            if tag == 'p':
                para = _extract_paragraph(child, idx)
                body_elements.append(para)
                if para['section_break']:
                    sections.append({'paragraph_index': idx, **para['section_break']})

            elif tag == 'sdt':
                # Structured Document Tag — preserve raw XML (breaking sdt causes
                # content-control corruption and incorrect formula numbering).
                body_elements.append({
                    'index': idx,
                    'type':  'raw_xml',
                    'xml':   ET.tostring(child, encoding='unicode'),
                })

            elif tag == 'tbl':
                body_elements.append(_extract_table(child, idx))

            elif tag == 'sectPr':
                # Final body-level section
                sections.append({'paragraph_index': None, **_extract_sectPr(child)})

            else:
                body_elements.append({'index': idx, 'type': tag})

            i += 1

        # 4. Headers and footers
        headers, footers = {}, {}
        for rId, rel in relationships.items():
            path = os.path.join(word_dir, rel['target'])
            if rel['type'] == 'header':
                headers[rel['target']] = _extract_hf_file(path)
            elif rel['type'] == 'footer':
                footers[rel['target']] = _extract_hf_file(path)

        # 5. Discover {{placeholder}} patterns in body text
        all_text = ' '.join(
            e['text'] for e in body_elements if e.get('type') == 'paragraph'
        )
        placeholders = sorted(set(re.findall(r'\{\{(\w+)\}\}', all_text)))

        # 6. Metadata
        metadata = _load_metadata(unzip_dir)

        result = {
            'source':        self.docx_path,
            'extracted_at':  datetime.now().isoformat(timespec='seconds'),
            'metadata':      metadata,
            'placeholders':  placeholders,
            'sections':      sections,
            'body_elements': body_elements,
            'headers':       headers,
            'footers':       footers,
            'relationships': relationships,
        }

        os.makedirs(output_dir, exist_ok=True)
        out_path = os.path.join(output_dir, 'extraction.json')
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        para_count  = sum(1 for e in body_elements if e.get('type') == 'paragraph')
        table_count = sum(1 for e in body_elements if e.get('type') == 'table')
        print(f'[extractor] Written → {out_path}')
        print(f'  paragraphs    : {para_count}')
        print(f'  tables        : {table_count}')
        print(f'  sections      : {len(sections)}')
        print(f'  headers       : {len(headers)}')
        print(f'  footers       : {len(footers)}')
        print(f'  placeholders  : {placeholders if placeholders else "(none found)"}')

        return result


if __name__ == '__main__':
    _docx    = sys.argv[1] if len(sys.argv) > 1 else 'template.docx'
    _out_dir = sys.argv[2] if len(sys.argv) > 2 else 'data'
    DocxExtractor(_docx).extract(output_dir=_out_dir)
