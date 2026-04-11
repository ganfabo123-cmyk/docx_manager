"""
Fills {{placeholder}} values in a .docx template and produces a new document.

Usage:
    python docx_restore.py [data.json] [output.docx]

Defaults:
    data.json       = data/data.json
    extraction.json = data/extraction.json   (produced by docx_extractor.py)
    output          = output.docx

Workflow:
  1. Load extraction.json (produced by docx_extractor.py) to learn which
     placeholders exist and which header/footer XML files are in use.
  2. Load data.json for the actual replacement values.
  3. Copy template/ to a temporary working directory (original is never modified).
  4. Apply substitutions to document.xml and only the header/footer files
     that are referenced in the extraction (no blind directory scans).
  5. Repack the result as a .docx ZIP archive.

Placeholder rules:
  - Syntax: {{field_name}}  (double curly braces, no spaces inside)
  - A placeholder may be split across multiple w:r runs by Word/WPS.
    The restorer merges run text before matching, then puts the result
    into the first run while preserving all run formatting (w:rPr).
"""

import base64 as _b64
import copy
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from xml.etree import ElementTree as ET

# ── Register all OOXML namespaces so ET preserves prefixes on serialization ───
_NS_MAP = {
    'wpc':          'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'mc':           'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':            'urn:schemas-microsoft-com:office:office',
    'r':            'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':            'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':            'urn:schemas-microsoft-com:vml',
    'wp14':         'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wp':           'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'w':            'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14':          'http://schemas.microsoft.com/office/word/2010/wordml',
    'w10':          'urn:schemas-microsoft-com:office:word',
    'w15':          'http://schemas.microsoft.com/office/word/2012/wordml',
    'wpg':          'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi':          'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wne':          'http://schemas.microsoft.com/office/word/2006/wordml',
    'wps':          'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpsCustomData':'http://www.wps.cn/officeDocument/2013/wpsCustomData',
    # DrawingML namespaces — must be registered so that anchor drawings,
    # inline images and WPS shapes keep standard prefixes when the XML is
    # re-serialised after placeholder substitution.  Without these, ET
    # generates ns0:/ns1: prefixes which can confuse WPS layout engine.
    'a':            'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':          'http://schemas.openxmlformats.org/drawingml/2006/picture',
}
for _prefix, _uri in _NS_MAP.items():
    ET.register_namespace(_prefix, _uri)

W    = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
WP_NS  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
MC_NS  = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
XML = 'http://www.w3.org/XML/1998/namespace'
XML_SPACE = f'{{{XML}}}space'

# WPS Office wraps these graphic URIs in mc:AlternateContent — must stay intact.
_WPS_URIS = frozenset({
    'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
})


def _q(local: str) -> str:
    return f'{{{W}}}{local}'


def _tag(elem: ET.Element) -> str:
    t = elem.tag
    return t.split('}', 1)[1] if '}' in t else t


# ── Core substitution logic ────────────────────────────────────────────────────

def _substitute_paragraph(p_elem: ET.Element, data: dict) -> bool:
    """
    Replace {{key}} placeholders in a single paragraph element.

    Strategy:
      - Collect only *text* runs — w:r children that have a w:t child.
        Runs that carry a w:drawing (anchor/inline images, shapes) or
        w:object (OLE formulas) have no w:t and are completely ignored.
        This is critical: the old code accidentally injected a w:t into
        drawing runs when they happened to be the first run in a paragraph
        that contained a placeholder, corrupting the anchor XML and causing
        floating shapes to render at wrong positions.
      - Merge text from all text runs; if no {{placeholder}} is present, exit.
      - Write the substituted text into the first text run's w:t and clear
        the rest.  Non-text runs (drawings, OLE) are never touched.

    Returns True if any substitution was made.
    """
    runs = [c for c in p_elem if _tag(c) == 'r']
    if not runs:
        return False

    # Only text runs (those with a w:t child) participate in substitution.
    # Drawing runs and OLE runs have no w:t and must not be modified.
    text_slots: list[tuple[ET.Element, ET.Element, str]] = []
    for r in runs:
        t = r.find(_q('t'))
        if t is not None:
            text_slots.append((r, t, t.text or ''))

    if not text_slots:
        return False

    full_text = ''.join(text for _, _, text in text_slots)
    if '{{' not in full_text:
        return False

    new_text = full_text
    for key, value in data.items():
        new_text = new_text.replace(f'{{{{{key}}}}}', str(value))

    if new_text == full_text:
        return False

    # Write replacement into the first text run
    first_r, first_t, _ = text_slots[0]
    first_t.text = new_text
    # xml:space="preserve" is required when text has leading/trailing whitespace
    if new_text and (new_text[0] == ' ' or new_text[-1] == ' '):
        first_t.set(XML_SPACE, 'preserve')
    else:
        first_t.attrib.pop(XML_SPACE, None)

    # Clear remaining text runs (drawing/OLE runs are untouched)
    for _, t, _ in text_slots[1:]:
        t.text = ''
        t.attrib.pop(XML_SPACE, None)

    return True


def _process_xml_file(xml_path: str, data: dict) -> int:
    """
    Parse xml_path, apply substitutions to all paragraphs, write back if changed.

    Returns the number of paragraphs modified.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    changed = 0
    for p in root.iter(_q('p')):
        if _substitute_paragraph(p, data):
            changed += 1

    if changed == 0:
        return 0

    xml_body = ET.tostring(root, encoding='unicode')
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        f.write(xml_body)

    return changed


# ── DOCX packing ───────────────────────────────────────────────────────────────

def _pack_docx(source_dir: str, output_path: str) -> None:
    """
    Zip source_dir contents into a .docx file at output_path.
    Files are stored with forward-slash paths (ZIP convention).
    Output is deterministic (sorted filenames).
    """
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for dirpath, dirnames, filenames in os.walk(source_dir):
            dirnames.sort()
            for filename in sorted(filenames):
                abs_path = os.path.join(dirpath, filename)
                arc_path = os.path.relpath(abs_path, source_dir).replace(os.sep, '/')
                zf.write(abs_path, arc_path)


# ── Compiler helper functions ──────────────────────────────────────────────────

def _parse_xml(xml_str: str) -> 'ET.Element | None':
    """Parse an XML string into an Element; return None on parse error."""
    try:
        return ET.fromstring(xml_str)
    except ET.ParseError:
        return None


def _write_xml(path: str, root: ET.Element) -> None:
    """Write root Element to path with an XML declaration."""
    xml_body = ET.tostring(root, encoding='unicode')
    with open(path, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        f.write(xml_body)


def _rels_append(rels_path: str, rid: str, type_uri: str, target: str) -> None:
    """Append a Relationship entry using string replacement (avoids ns loss)."""
    with open(rels_path, 'r', encoding='utf-8') as f:
        content = f.read()
    tag = f'<Relationship Id="{rid}" Type="{type_uri}" Target="{target}"/>'
    content = content.replace('</Relationships>', f'  {tag}\n</Relationships>')
    with open(rels_path, 'w', encoding='utf-8') as f:
        f.write(content)


def _ensure_content_type(ct_path: str, ext: str, content_type: str) -> None:
    """Add a Default content-type entry if not already present (string op)."""
    with open(ct_path, 'r', encoding='utf-8') as f:
        content = f.read()
    if f'Extension="{ext}"' in content:
        return
    tag = f'<Default Extension="{ext}" ContentType="{content_type}"/>'
    content = content.replace('</Types>', f'  {tag}\n</Types>')
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(content)


def _image_suffix(data: bytes) -> str:
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        return 'png'
    if data[:3] == b'\xff\xd8\xff':
        return 'jpg'
    if data[:4] in (b'II*\x00', b'MM\x00*'):
        return 'tiff'
    if data[:4] == b'\xd7\xcd\xc6\x9a':
        return 'wmf'
    if data[:6] in (b'GIF87a', b'GIF89a'):
        return 'gif'
    return 'png'


_IMAGE_CT = {
    'png':  'image/png',
    'jpg':  'image/jpeg',
    'jpeg': 'image/jpeg',
    'tiff': 'image/tiff',
    'wmf':  'image/x-wmf',
    'gif':  'image/gif',
}


def _minimal_tbl() -> ET.Element:
    tbl  = ET.Element(_q('tbl'))
    tblPr = ET.SubElement(tbl, _q('tblPr'))
    ts = ET.SubElement(tblPr, _q('tblStyle'))
    ts.set(_q('val'), 'TableGrid')
    tw = ET.SubElement(tblPr, _q('tblW'))
    tw.set(_q('type'), 'auto')
    return tbl


def _minimal_tc() -> ET.Element:
    tc = ET.Element(_q('tc'))
    ET.SubElement(tc, _q('tcPr'))
    return tc


def _plain_para(text: str) -> ET.Element:
    p = ET.Element(_q('p'))
    if text:
        r = ET.SubElement(p, _q('r'))
        t = ET.SubElement(r, _q('t'))
        t.text = text
        if text[0] == ' ' or text[-1] == ' ':
            t.set(XML_SPACE, 'preserve')
    return p


def _is_wps_anchor_drawing(elem: ET.Element) -> bool:
    """Return True if <w:drawing> contains a WPS-specific anchor graphic."""
    if _tag(elem) != 'drawing':
        return False
    for anchor in elem:
        if _tag(anchor) == 'anchor':
            for gd in anchor.iter(f'{{{A_NS}}}graphicData'):
                if gd.get('uri', '') in _WPS_URIS:
                    return True
    return False


def _wrap_in_mc_choice(drawing_elem: ET.Element) -> ET.Element:
    """Wrap a bare <w:drawing> in mc:AlternateContent/mc:Choice[Requires='wps']."""
    alt    = ET.Element(f'{{{MC_NS}}}AlternateContent')
    choice = ET.SubElement(alt, f'{{{MC_NS}}}Choice')
    choice.set('Requires', 'wps')
    choice.append(drawing_elem)
    return alt


def _math_text(oMath: ET.Element, text: str) -> None:
    """Append a plain-text run to an <m:oMath> element."""
    r = ET.SubElement(oMath, f'{{{M_NS}}}r')
    t = ET.SubElement(r, f'{{{M_NS}}}t')
    t.text = text


# ── DocxCompiler ───────────────────────────────────────────────────────────────

class DocxCompiler:
    """
    Rebuilds a .docx from extraction.json + template directory.

    Unlike DocxRestorer (which performs in-place placeholder substitution),
    DocxCompiler completely rewrites document.xml and all header/footer files
    from the structured data in extraction.json, while copying binary assets
    (media/, embeddings/, styles.xml, etc.) unchanged from the template.

    Usage:
        DocxCompiler(
            extraction_path = 'data/extraction.json',
            template_dir    = 'template',
            data            = {'field_a': 'value', …},
        ).compile('output.docx')
    """

    def __init__(
        self,
        extraction_path: str        = 'data/extraction.json',
        template_dir:    str        = 'template',
        data:            'dict | None' = None,
    ):
        if not os.path.exists(extraction_path):
            raise FileNotFoundError(
                f'[compiler] extraction.json not found: {extraction_path!r}. '
                'Run docx_extractor.py first.'
            )
        with open(extraction_path, 'r', encoding='utf-8') as f:
            self.ext = json.load(f)
        self.extraction_path = extraction_path
        self.template_dir    = template_dir
        self.data            = data or {}
        self._rid_counter    = 1
        self._shape_counter  = 1

    # ── rId allocation ─────────────────────────────────────────────────────────

    def _init_rid_counter(self, work_dir: str) -> None:
        """Set _rid_counter to max(existing rId numbers) + 1."""
        rels_path = os.path.join(work_dir, 'word', '_rels', 'document.xml.rels')
        if not os.path.exists(rels_path):
            self._rid_counter = 1
            return
        with open(rels_path, encoding='utf-8') as f:
            content = f.read()
        nums = [int(x) for x in re.findall(r'Id="rId(\d+)"', content)]
        self._rid_counter = max(nums, default=0) + 1

    def _alloc_rid(self) -> str:
        rid = f'rId{self._rid_counter}'
        self._rid_counter += 1
        return rid

    # ── Paragraph / run builder ────────────────────────────────────────────────

    def _build_para(self, pdata: dict) -> ET.Element:
        """Reconstruct a <w:p> from an extraction paragraph dict."""
        p = ET.Element(_q('p'))

        pPr_str = pdata.get('pPr')
        if pPr_str:
            pPr_elem = _parse_xml(pPr_str)
            if pPr_elem is not None:
                p.append(pPr_elem)

        for run in pdata.get('runs', []):
            r = ET.Element(_q('r'))

            rPr_str = run.get('rPr')
            if rPr_str:
                rPr_elem = _parse_xml(rPr_str)
                if rPr_elem is not None:
                    r.append(rPr_elem)

            drawing_xml = run.get('drawing_xml')
            object_xml  = run.get('object_xml')

            if drawing_xml:
                node = _parse_xml(drawing_xml)
                if node is not None:
                    if _tag(node) == 'AlternateContent':
                        # WPS anchor with mc:Fallback VML — must stay intact
                        r.append(node)
                    elif _tag(node) == 'drawing' and _is_wps_anchor_drawing(node):
                        # Old-format bare drawing — wrap for WPS compatibility
                        r.append(_wrap_in_mc_choice(copy.deepcopy(node)))
                    else:
                        r.append(node)

            elif object_xml:
                node = _parse_xml(object_xml)
                if node is not None:
                    r.append(node)

            else:
                text = run.get('text', '') or ''
                if self.data:
                    for key, val in self.data.items():
                        text = text.replace(f'{{{{{key}}}}}', str(val))
                t = ET.SubElement(r, _q('t'))
                t.text = text
                if text and (text[0] == ' ' or text[-1] == ' '):
                    t.set(XML_SPACE, 'preserve')
                else:
                    t.attrib.pop(XML_SPACE, None)

            p.append(r)

        return p

    # ── Specialised element builders ───────────────────────────────────────────

    def _build_toc_para(self, elem: dict) -> ET.Element:
        """Generate a TOC field paragraph (w:dirty forces re-calculation on open)."""
        max_level = elem.get('max_level', 4)
        p = ET.Element(_q('p'))

        r1 = ET.SubElement(p, _q('r'))
        fc1 = ET.SubElement(r1, _q('fldChar'))
        fc1.set(_q('fldCharType'), 'begin')
        fc1.set(_q('dirty'), 'true')

        r2 = ET.SubElement(p, _q('r'))
        instr = ET.SubElement(r2, _q('instrText'))
        instr.set(XML_SPACE, 'preserve')
        instr.text = f' TOC \\o "1-{max_level}" \\z '

        r3 = ET.SubElement(p, _q('r'))
        ET.SubElement(r3, _q('fldChar')).set(_q('fldCharType'), 'separate')

        r4 = ET.SubElement(p, _q('r'))
        ET.SubElement(r4, _q('fldChar')).set(_q('fldCharType'), 'end')

        return p

    def _build_table(
        self,
        tdata:    dict,
        tmpl_tbl: 'ET.Element | None',
    ) -> ET.Element:
        """Reconstruct <w:tbl>: content from extraction, structure from template."""
        if tmpl_tbl is not None:
            tbl = copy.deepcopy(tmpl_tbl)
            for tr in list(tbl.findall(_q('tr'))):
                tbl.remove(tr)
            tmpl_rows = tmpl_tbl.findall(_q('tr'))
        else:
            tbl       = _minimal_tbl()
            tmpl_rows = []

        for row_idx, row_data in enumerate(tdata.get('rows', [])):
            if tmpl_rows:
                tmpl_row   = copy.deepcopy(tmpl_rows[min(row_idx, len(tmpl_rows) - 1)])
                tmpl_cells = tmpl_row.findall(_q('tc'))
                for tc in list(tmpl_row.findall(_q('tc'))):
                    tmpl_row.remove(tc)
            else:
                tmpl_row   = ET.Element(_q('tr'))
                tmpl_cells = []

            for cell_idx, cell_data in enumerate(row_data):
                if tmpl_cells:
                    tc = copy.deepcopy(tmpl_cells[min(cell_idx, len(tmpl_cells) - 1)])
                    for p in list(tc.findall(_q('p'))):
                        tc.remove(p)
                else:
                    tc = _minimal_tc()

                for para_data in cell_data.get('paragraphs', []):
                    tc.append(self._build_para(para_data))

                if not tc.findall(_q('p')):
                    tc.append(_plain_para(''))

                tmpl_row.append(tc)

            tbl.append(tmpl_row)

        return tbl

    def _build_image_nodes(self, elem: dict, work_dir: str) -> list:
        """Return one or more <w:p> elements for a type='image' body element."""
        caption = elem.get('caption', '')
        fallback_label = f'[图片: {caption}]' if caption else '[图片]'

        # Path A: drawing_xml already serialised (preserved from template)
        if 'drawing_xml' in elem:
            node = _parse_xml(elem['drawing_xml'])
            if node is None:
                return [_plain_para(fallback_label)]
            p = ET.Element(_q('p'))
            r = ET.SubElement(p, _q('r'))
            r.append(node)
            nodes = [p]
            if caption:
                nodes.append(_plain_para(caption))
            return nodes

        # Path B: base64-encoded image bytes
        b64 = elem.get('base64', '')
        if not b64:
            return [_plain_para(fallback_label)]

        try:
            img_bytes = _b64.b64decode(b64)
        except Exception:
            return [_plain_para(fallback_label)]

        suffix    = _image_suffix(img_bytes)
        media_dir = os.path.join(work_dir, 'word', 'media')
        os.makedirs(media_dir, exist_ok=True)

        img_num  = len([f for f in os.listdir(media_dir) if f.startswith('image')]) + 1
        img_name = f'image{img_num}.{suffix}'
        with open(os.path.join(media_dir, img_name), 'wb') as f:
            f.write(img_bytes)

        rid = self._alloc_rid()
        _rels_append(
            os.path.join(work_dir, 'word', '_rels', 'document.xml.rels'),
            rid,
            f'{R_NS}/image',
            f'media/{img_name}',
        )
        _ensure_content_type(
            os.path.join(work_dir, '[Content_Types].xml'),
            suffix,
            _IMAGE_CT.get(suffix, 'image/png'),
        )

        width_pt  = elem.get('width',  120) or 120
        height_pt = elem.get('height', 80)  or 80
        cx = int(width_pt  * 12700)   # pt → EMU
        cy = int(height_pt * 12700)

        shape_id = self._shape_counter
        self._shape_counter += 1

        p    = ET.Element(_q('p'))
        pPr  = ET.SubElement(p, _q('pPr'))
        jc   = ET.SubElement(pPr, _q('jc'))
        jc.set(_q('val'), elem.get('position', 'center'))

        r       = ET.SubElement(p, _q('r'))
        drawing = ET.SubElement(r, _q('drawing'))
        inline  = ET.SubElement(drawing, f'{{{WP_NS}}}inline')
        for k, v in (('distT', '0'), ('distB', '0'), ('distL', '0'), ('distR', '0')):
            inline.set(k, v)

        ext = ET.SubElement(inline, f'{{{WP_NS}}}extent')
        ext.set('cx', str(cx)); ext.set('cy', str(cy))

        ee = ET.SubElement(inline, f'{{{WP_NS}}}effectExtent')
        for k in ('l', 't', 'r', 'b'):
            ee.set(k, '0')

        docPr = ET.SubElement(inline, f'{{{WP_NS}}}docPr')
        docPr.set('id', str(shape_id))
        docPr.set('name', f'Picture {shape_id}')

        cNvFr = ET.SubElement(inline, f'{{{WP_NS}}}cNvGraphicFramePr')
        locks = ET.SubElement(cNvFr, f'{{{A_NS}}}graphicFrameLocks')
        locks.set('noChangeAspect', '1')

        graphic     = ET.SubElement(inline, f'{{{A_NS}}}graphic')
        graphicData = ET.SubElement(graphic, f'{{{A_NS}}}graphicData')
        graphicData.set('uri', PIC_NS)

        pic = ET.SubElement(graphicData, f'{{{PIC_NS}}}pic')

        nvPicPr = ET.SubElement(pic, f'{{{PIC_NS}}}nvPicPr')
        cNvPr   = ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPr')
        cNvPr.set('id', str(shape_id))
        cNvPr.set('name', img_name)
        ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPicPr')

        blipFill = ET.SubElement(pic, f'{{{PIC_NS}}}blipFill')
        blip     = ET.SubElement(blipFill, f'{{{A_NS}}}blip')
        blip.set(f'{{{R_NS}}}embed', rid)
        stretch  = ET.SubElement(blipFill, f'{{{A_NS}}}stretch')
        ET.SubElement(stretch, f'{{{A_NS}}}fillRect')

        spPr    = ET.SubElement(pic, f'{{{PIC_NS}}}spPr')
        xfrm    = ET.SubElement(spPr, f'{{{A_NS}}}xfrm')
        off     = ET.SubElement(xfrm, f'{{{A_NS}}}off')
        off.set('x', '0'); off.set('y', '0')
        ext2    = ET.SubElement(xfrm, f'{{{A_NS}}}ext')
        ext2.set('cx', str(cx)); ext2.set('cy', str(cy))
        prstGeom = ET.SubElement(spPr, f'{{{A_NS}}}prstGeom')
        prstGeom.set('prst', 'rect')
        ET.SubElement(prstGeom, f'{{{A_NS}}}avLst')

        nodes = [p]
        if caption:
            nodes.append(_plain_para(caption))
        return nodes

    def _build_ole_para(self, elem: dict, work_dir: str) -> ET.Element:
        """Build a paragraph containing an OLE embedded object (Equation Editor)."""
        V_NS = 'urn:schemas-microsoft-com:vml'
        O_NS = 'urn:schemas-microsoft-com:office:office'

        p   = ET.Element(_q('p'))
        pPr = ET.SubElement(p, _q('pPr'))
        jc  = ET.SubElement(pPr, _q('jc'))
        jc.set(_q('val'), 'center')

        b64 = elem.get('base64', '')
        if b64:
            try:
                ole_bytes = _b64.b64decode(b64)
                emb_dir   = os.path.join(work_dir, 'word', 'embeddings')
                os.makedirs(emb_dir, exist_ok=True)
                ole_num  = len([f for f in os.listdir(emb_dir)
                                if f.startswith('oleObject')]) + 1
                ole_name = f'oleObject{ole_num}.bin'
                with open(os.path.join(emb_dir, ole_name), 'wb') as f:
                    f.write(ole_bytes)

                rid = self._alloc_rid()
                _rels_append(
                    os.path.join(work_dir, 'word', '_rels', 'document.xml.rels'),
                    rid,
                    f'{R_NS}/oleObject',
                    f'embeddings/{ole_name}',
                )

                shape_id = self._shape_counter
                self._shape_counter += 1
                sid_str  = f'_x0000_i{1024 + shape_id}'

                r   = ET.SubElement(p, _q('r'))
                obj = ET.SubElement(r, _q('object'))
                obj.set(_q('dxaOrig'), '2400')
                obj.set(_q('dyaOrig'), '600')

                shape = ET.SubElement(obj, f'{{{V_NS}}}shape')
                shape.set('id',     sid_str)
                shape.set('style',  'width:120pt;height:30pt')
                shape.set(f'{{{O_NS}}}ole', '')

                ole_obj = ET.SubElement(obj, f'{{{O_NS}}}OLEObject')
                ole_obj.set('Type',        'Embed')
                ole_obj.set('ProgID',      'Equation.3')
                ole_obj.set(f'{{{R_NS}}}id', rid)
                ole_obj.set('ShapeID',     sid_str)
                ole_obj.set('DrawAspect',  'Content')

            except Exception:
                p.append(_plain_para('[公式]'))

        formula_index = elem.get('formula_index', '')
        if formula_index:
            r_idx = ET.SubElement(p, _q('r'))
            t     = ET.SubElement(r_idx, _q('t'))
            t.text = f'  {formula_index}'
            t.set(XML_SPACE, 'preserve')

        return p

    def _build_omath_para(self, elem: dict) -> ET.Element:
        """Build an OMML math formula paragraph."""
        formula       = elem.get('formula', '')
        formula_index = elem.get('formula_index', '')

        p         = ET.Element(_q('p'))
        oMathPara = ET.SubElement(p, f'{{{M_NS}}}oMathPara')
        oMath     = ET.SubElement(oMathPara, f'{{{M_NS}}}oMath')

        if formula.startswith('<'):
            node = _parse_xml(formula)
            if node is not None:
                if _tag(node) == 'oMath':
                    for child in list(node):
                        oMath.append(child)
                else:
                    oMath.append(node)
            else:
                _math_text(oMath, formula)
        else:
            _math_text(oMath, formula)

        if formula_index:
            r = ET.SubElement(p, _q('r'))
            t = ET.SubElement(r, _q('t'))
            t.text = f'  {formula_index}'
            t.set(XML_SPACE, 'preserve')

        return p

    # ── Document rebuild ───────────────────────────────────────────────────────

    def _rebuild_document(self, work_dir: str) -> None:
        """Completely rewrite word/document.xml from extraction.json body_elements."""
        doc_path = os.path.join(work_dir, 'word', 'document.xml')
        tree     = ET.parse(doc_path)
        root     = tree.getroot()
        body     = root.find(_q('body'))

        # Collect template tables (kth table in extraction ↔ kth template table)
        tmpl_tables = [c for c in body if _tag(c) == 'tbl']

        # Preserve body-level sectPr (page margins, paper size, etc.)
        tmpl_sectPr = body.find(_q('sectPr'))
        if tmpl_sectPr is not None:
            tmpl_sectPr = copy.deepcopy(tmpl_sectPr)

        # Clear body
        for child in list(body):
            body.remove(child)

        tbl_idx = 0
        for elem in self.ext.get('body_elements', []):
            etype = elem.get('type')

            if etype == 'paragraph':
                body.append(self._build_para(elem))

            elif etype == 'table':
                tmpl_tbl = tmpl_tables[tbl_idx] if tbl_idx < len(tmpl_tables) else None
                body.append(self._build_table(elem, tmpl_tbl))
                tbl_idx += 1

            elif etype == 'raw_xml':
                # TOC, fldChar blocks, sdt — emit verbatim
                node = _parse_xml(elem.get('xml', ''))
                if node is not None:
                    body.append(node)

            elif etype == 'toc':
                body.append(self._build_toc_para(elem))

            elif etype == 'image':
                for node in self._build_image_nodes(elem, work_dir):
                    body.append(node)

            elif etype == 'ole':
                body.append(self._build_ole_para(elem, work_dir))

            elif etype in ('omath', 'omathpara'):
                body.append(self._build_omath_para(elem))

            # bookmarkEnd, bookmarkStart, and other internal markers are skipped

        # Restore body-level sectPr at the very end
        if tmpl_sectPr is not None:
            body.append(tmpl_sectPr)
        elif not list(body):
            body.append(_plain_para(''))

        _write_xml(doc_path, root)

    def _rebuild_hf(self, xml_path: str, hf_data: dict) -> None:
        """Rewrite a single header/footer XML file from extraction data."""
        if not os.path.exists(xml_path):
            return
        tree = ET.parse(xml_path)
        root = tree.getroot()

        for child in list(root):
            root.remove(child)

        for pdata in hf_data.get('paragraphs', []):
            root.append(self._build_para(pdata))

        # OOXML requires at least one paragraph per hdr/ftr element
        if not list(root):
            root.append(_plain_para(''))

        _write_xml(xml_path, root)

    def _rebuild_hf_files(self, work_dir: str) -> None:
        """Rebuild all header and footer XML files."""
        word_dir = os.path.join(work_dir, 'word')
        for filename, hf_data in self.ext.get('headers', {}).items():
            if hf_data:
                self._rebuild_hf(os.path.join(word_dir, filename), hf_data)
        for filename, hf_data in self.ext.get('footers', {}).items():
            if hf_data:
                self._rebuild_hf(os.path.join(word_dir, filename), hf_data)

    def _patch_settings(self, work_dir: str) -> None:
        """Inject <w:updateFields w:val="true"/> so Word/WPS recalculates TOC on open."""
        settings_path = os.path.join(work_dir, 'word', 'settings.xml')
        if not os.path.exists(settings_path):
            return
        with open(settings_path, 'r', encoding='utf-8') as f:
            content = f.read()
        if 'updateFields' in content:
            return
        content = content.replace(
            '</w:settings>',
            '  <w:updateFields w:val="true"/>\n</w:settings>',
        )
        with open(settings_path, 'w', encoding='utf-8') as f:
            f.write(content)

    # ── Main compile entry point ───────────────────────────────────────────────

    def compile(self, output_path: str = 'output.docx') -> str:
        """
        Produce a filled .docx at output_path by fully rebuilding from
        extraction.json.  Placeholder substitution ({{key}}) is applied
        during _build_para() for all text runs.

        Returns the absolute path to the generated file.
        """
        bodies    = self.ext.get('body_elements', [])
        raw_count  = sum(1 for e in bodies if e.get('type') == 'raw_xml')
        para_count = sum(1 for e in bodies if e.get('type') == 'paragraph')
        tbl_count  = sum(1 for e in bodies if e.get('type') == 'table')
        print(f'[compiler] Loaded extraction from {self.extraction_path!r}')
        print(f'           body_elements : {len(bodies)}'
              f' ({para_count} para, {tbl_count} tbl, {raw_count} raw_xml)')

        tmp_base = tempfile.mkdtemp(prefix='docx_compile_')
        work_dir = os.path.join(tmp_base, 'work')
        try:
            shutil.copytree(self.template_dir, work_dir)
            self._init_rid_counter(work_dir)
            self._rebuild_document(work_dir)
            self._rebuild_hf_files(work_dir)
            self._patch_settings(work_dir)
            _pack_docx(work_dir, output_path)
        finally:
            shutil.rmtree(tmp_base, ignore_errors=True)

        abs_out = os.path.abspath(output_path)
        print(f'[compiler] Done → {abs_out}')
        return abs_out


# ── Main restorer class ────────────────────────────────────────────────────────

class DocxRestorer:
    def __init__(
        self,
        template_dir:    str        = 'template',
        data:            dict | None = None,
        extraction_path: str        = 'data/extraction.json',
    ):
        """
        Args:
            template_dir:    path to the extracted (unzipped) template directory.
            data:            dict mapping placeholder names to replacement values.
            extraction_path: path to extraction.json produced by DocxExtractor.
        """
        self.template_dir    = template_dir
        self.data            = data or {}
        self.extraction_path = extraction_path

    # ── Load and validate against extraction ──────────────────────────────────

    def _load_extraction(self) -> dict:
        """
        Load extraction.json and return the parsed dict.
        Raises FileNotFoundError if the file does not exist.
        """
        if not os.path.exists(self.extraction_path):
            raise FileNotFoundError(
                f'[restorer] extraction.json not found at {self.extraction_path!r}. '
                'Run docx_extractor.py first.'
            )
        with open(self.extraction_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _validate_data(self, known_placeholders: list[str]) -> None:
        """
        Compare self.data keys against the placeholders found during extraction.
        Prints warnings for missing values and unused keys.
        """
        known = set(known_placeholders)
        provided = set(self.data.keys())

        missing = known - provided
        unused  = provided - known

        if missing:
            print(f'[restorer] Warning: {len(missing)} placeholder(s) in template have no data value:')
            for k in sorted(missing):
                print(f'             missing → {{{{{k}}}}}')

        if unused:
            print(f'[restorer] Warning: {len(unused)} data key(s) do not match any placeholder:')
            for k in sorted(unused):
                print(f'             unused  → {k}')

    # ── Restore ────────────────────────────────────────────────────────────────

    def restore(self, output_path: str = 'output.docx') -> str:
        """
        Produce a filled .docx at output_path.

        Steps:
          1. Parse extraction.json for the list of placeholders and the set of
             header/footer filenames that are actually referenced in this template.
          2. Validate self.data keys against those placeholders and warn on mismatches.
          3. Work on a temporary copy of template/ — the original is never modified.
          4. Substitute placeholders in document.xml and each referenced
             header/footer file.
          5. Repack as a .docx ZIP archive.

        Returns:
            Absolute path to the generated file.
        """
        # 1. Load extraction.json
        extraction = self._load_extraction()
        known_placeholders: list[str] = extraction.get('placeholders', [])
        hf_filenames: list[str] = sorted(
            set(extraction.get('headers', {}).keys()) |
            set(extraction.get('footers', {}).keys())
        )

        print(f'[restorer] Loaded extraction from {self.extraction_path!r}')
        print(f'           placeholders : {known_placeholders if known_placeholders else "(none)"}')
        print(f'           header/footer files : {hf_filenames if hf_filenames else "(none)"}')

        # 2. Validate data keys
        if not self.data:
            print('[restorer] Warning: no data provided — output will be identical to template.')
        else:
            self._validate_data(known_placeholders)

        # 3. Work on a temporary copy so the original template/ is always preserved
        tmp_base = tempfile.mkdtemp(prefix='docx_restore_')
        work_dir = os.path.join(tmp_base, 'work')
        try:
            shutil.copytree(self.template_dir, work_dir)
            word_dir = os.path.join(work_dir, 'word')

            total_changed = 0

            # 4a. Process main document body
            doc_path = os.path.join(word_dir, 'document.xml')
            n = _process_xml_file(doc_path, self.data)
            if n:
                print(f'[restorer]   document.xml          : {n} paragraph(s) modified')
            total_changed += n

            # 4b. Process only the header/footer files referenced in the extraction
            for filename in hf_filenames:
                xml_path = os.path.join(word_dir, filename)
                if not os.path.exists(xml_path):
                    print(f'[restorer]   {filename}: not found in template, skipping')
                    continue
                n = _process_xml_file(xml_path, self.data)
                if n:
                    print(f'[restorer]   {filename:<22}: {n} paragraph(s) modified')
                total_changed += n

            if total_changed == 0:
                print('[restorer] No placeholders matched — check your data keys.')

            # 5. Repack
            _pack_docx(work_dir, output_path)

        finally:
            shutil.rmtree(tmp_base, ignore_errors=True)

        abs_out = os.path.abspath(output_path)
        print(f'[restorer] Done → {abs_out}')
        return abs_out


if __name__ == '__main__':
    _data_path       = sys.argv[1] if len(sys.argv) > 1 else 'data/data.json'
    _output_path     = sys.argv[2] if len(sys.argv) > 2 else 'output.docx'
    _extraction_path = sys.argv[3] if len(sys.argv) > 3 else 'data/extraction.json'

    if os.path.exists(_data_path):
        with open(_data_path, 'r', encoding='utf-8') as _f:
            _data = json.load(_f)
        print(f'[compiler] Loaded {len(_data)} field(s) from {_data_path!r}')
    else:
        print(f'[compiler] Data file not found: {_data_path!r} — running with empty substitutions.')
        _data = {}

    DocxCompiler(
        extraction_path=_extraction_path,
        template_dir='template',
        data=_data,
    ).compile(output_path=_output_path)
