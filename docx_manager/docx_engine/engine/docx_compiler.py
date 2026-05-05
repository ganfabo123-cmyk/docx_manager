"""
Compiles a .docx from extraction.json + template directory.

Primary source : extraction.json  — body_elements, headers, footers, metadata
Supporting role: template/        — table structure (tblPr/tblGrid/trPr/tcPr),
                                    body-level sectPr, binary assets, styles,
                                    settings, theme, relationships, content types

The template is consulted ONLY for data that extraction.json does not store:
  • <w:tblPr> / <w:tblGrid> / <w:trPr> / <w:tcPr>  (table structural XML)
  • The body-level <w:sectPr>  (page size, margins, section layout)
  • The document root element  (namespace declarations are reused)
  • Binary blobs: media/, embeddings/  (never modified)
  • styles.xml, settings.xml, theme/   (never modified)

Every paragraph (pPr, runs, rPr) and every header/footer paragraph comes
exclusively from extraction.json.

Usage:
    python docx_compiler.py [extraction.json] [output.docx] [template_dir]

Defaults:
    extraction  = data/extraction.json
    output      = output.docx
    template    = template
"""

import base64
import copy
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path
from pydantic import BaseModel

try:
    from . import base_agent as ba
except ImportError:
    import base_agent as ba  # type: ignore

# ── Namespace registration ─────────────────────────────────────────────────────
# Must happen before any ET.parse() / ET.tostring() call so prefixes are stable.

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
    # DrawingML namespaces (needed for inline images)
    'a':            'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':          'http://schemas.openxmlformats.org/drawingml/2006/picture',
}
for _p, _u in _NS_MAP.items():
    ET.register_namespace(_p, _u)

# Namespace URI constants
W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M   = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT  = 'http://schemas.openxmlformats.org/package/2006/content-types'
A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

REL_TYPE_IMAGE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
)
REL_TYPE_OLE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'
)

# VML and Office namespace URIs (already registered in _NS_MAP; constants
# are needed when constructing elements programmatically).
V_NS = 'urn:schemas-microsoft-com:vml'
O_NS = 'urn:schemas-microsoft-com:office:office'
MC   = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

# HIT A4 page text-area width in twips (210mm − 2×25.4mm margins ≈ 159.2mm → 9072twips,
# minus tblInd 138 twips used by the templates → effective ≈ 8296 twips).
_TABLE_TEXT_WIDTH_TWIPS = 8296


def _q(local: str) -> str:   return f'{{{W}}}{local}'
def _qm(local: str) -> str:  return f'{{{M}}}{local}'
def _qr(local: str) -> str:  return f'{{{R}}}{local}'

def _tag(elem: ET.Element) -> str:
    t = elem.tag
    return t.split('}', 1)[1] if '}' in t else t


# ── XML fragment parser ────────────────────────────────────────────────────────

def _parse_xml(xml_str: str | None) -> ET.Element | None:
    """
    Parse an XML fragment string from extraction.json into an Element.

    Extraction stores pPr/rPr with the ns0: prefix (an artefact of ET
    serialising without registered prefixes in the extractor).  fromstring()
    resolves ns0: to the correct URI; tostring() then re-serialises with
    the registered w: prefix.  No manual prefix rewriting is needed.
    """
    if not xml_str:
        return None
    try:
        return ET.fromstring(xml_str)
    except ET.ParseError:
        return None


# ── DOCX packing ──────────────────────────────────────────────────────────────

def _pack_docx(source_dir: str, output_path: str) -> None:
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for dirpath, dirnames, filenames in os.walk(source_dir):
            dirnames.sort()
            for filename in sorted(filenames):
                abs_path = os.path.join(dirpath, filename)
                arc_path = os.path.relpath(abs_path, source_dir).replace(os.sep, '/')
                zf.write(abs_path, arc_path)


# ── XML writer ─────────────────────────────────────────────────────────────────

def _write_xml(path: str, root: ET.Element) -> None:
    with open(path, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        f.write(ET.tostring(root, encoding='unicode'))


# ── Image format detection ─────────────────────────────────────────────────────

def _image_suffix(data: bytes) -> str:
    if data[:4]  == b'\x89PNG':                        return '.png'
    if data[:2]  == b'\xff\xd8':                       return '.jpg'
    if data[:4]  in (b'II\x2a\x00', b'MM\x00\x2a'):   return '.tiff'
    if data[:4]  == b'\xd7\xcd\xc6\x9a':              return '.wmf'
    if data[:6]  in (b'GIF87a', b'GIF89a'):            return '.gif'
    return '.png'


# ── .rels file helper (string-based to avoid default-namespace issues) ─────────

def _rels_append(rels_path: str, rid: str, rel_type: str, target: str) -> None:
    """Insert a new Relationship element into a .rels file."""
    with open(rels_path, 'r', encoding='utf-8') as f:
        text = f.read()
    entry = (
        f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/>'
    )
    close = '</Relationships>'
    text = text.replace(close, entry + close)
    with open(rels_path, 'w', encoding='utf-8') as f:
        f.write(text)


# ── [Content_Types].xml helper ─────────────────────────────────────────────────

_MIME: dict[str, str] = {
    'png':  'image/png',
    'jpg':  'image/jpeg',
    'jpeg': 'image/jpeg',
    'tif':  'image/tiff',
    'tiff': 'image/tiff',
    'wmf':  'image/x-wmf',
    'gif':  'image/gif',
    'bin':  'application/vnd.openxmlformats-officedocument.oleObject',
}

def _ensure_content_type(ct_path: str, ext: str) -> None:
    """Add a Default entry for ``ext`` if absent (string-based)."""
    ext = ext.lower().lstrip('.')
    with open(ct_path, 'r', encoding='utf-8') as f:
        text = f.read()
    if f'Extension="{ext}"' in text:
        return
    mime  = _MIME.get(ext, 'application/octet-stream')
    entry = f'<Default Extension="{ext}" ContentType="{mime}"/>'
    close = '</Types>'
    text  = text.replace(close, entry + close)
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(text)


# ── Anchor-drawing helpers ─────────────────────────────────────────────────────

# WPS-specific graphicData URIs that require mc:AlternateContent wrapping.
_WPS_URIS = {
    'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
}


def _is_wps_anchor_drawing(elem: ET.Element) -> bool:
    """
    Return True when elem is a bare <w:drawing> containing a wp:anchor whose
    graphicData URI points to a WPS-specific shape namespace.

    These drawings MUST be wrapped in <mc:AlternateContent><mc:Choice
    Requires="wps"> so that WPS Office honours the anchor's posOffset
    values.  Without the wrapper WPS ignores the stored position and renders
    the shape at an unpredictable location.
    """
    if _tag(elem) != 'drawing':
        return False
    anchor = elem.find(f'{{{WP}}}anchor')
    if anchor is None:
        return False
    for gd in anchor.iter(f'{{{A}}}graphicData'):
        if gd.get('uri', '') in _WPS_URIS:
            return True
    return False


def _wrap_in_mc_choice(drawing_elem: ET.Element) -> ET.Element:
    """
    Wrap a <w:drawing> in <mc:AlternateContent><mc:Choice Requires="wps">.

    Used by the compiler to restore the mc:AlternateContent structure that
    the extractor stripped when it stored only the inner w:drawing from an
    mc:Choice block.  The mc:Fallback (VML) section is omitted because the
    extractor did not preserve it; WPS will still render correctly using the
    Choice path.
    """
    alt    = ET.Element(f'{{{MC}}}AlternateContent')
    choice = ET.SubElement(alt, f'{{{MC}}}Choice')
    choice.set('Requires', 'wps')
    choice.append(drawing_elem)
    return alt


# ── Abstract heading detection result ─────────────────────────────────────────

class AbstractCheckResult(BaseModel):
    has_abstract_cn: bool  # 是否存在"摘要"一级标题
    has_abstract_en: bool  # 是否存在"Abstract"一级标题


# ── Main compiler ──────────────────────────────────────────────────────────────

class DocxCompiler:
    """
    Compile a .docx from extraction.json and a template directory.

    The template is used ONLY to provide structural data not stored in
    extraction.json: table skeleton XML, body-level sectPr, binary assets,
    and the document root element (for namespace declarations).

    Every content decision — paragraph text, run formatting, pPr, rPr,
    header and footer text — is driven by extraction.json.
    """

    def __init__(
        self,
        extraction_path: str = 'data/extraction.json',
        template_dir:    str = 'template',
    ) -> None:
        if not os.path.exists(extraction_path):
            raise FileNotFoundError(f'extraction not found: {extraction_path!r}')
        if not os.path.isdir(template_dir):
            raise FileNotFoundError(f'template dir not found: {template_dir!r}')

        with open(extraction_path, 'r', encoding='utf-8') as f:
            self.ext: dict = json.load(f)

        self.template_dir    = template_dir
        self.extraction_path = extraction_path
        self._rid_counter    = 0   # next unused rId number
        self._shape_counter  = 1   # docPr id for inline images
        self.deferred_images: list[dict] = []
        self.abstract_check: AbstractCheckResult = AbstractCheckResult(
            has_abstract_cn=False, has_abstract_en=False
        )

    # ── Public API ─────────────────────────────────────────────────────────────

    def compile(
        self,
        output_path: str = 'output.docx',
        skip_images: bool = True,
    ) -> str:
        """
        Build the .docx and write it to output_path.

        skip_images:
            When True, image elements are NOT written into the document XML.
            Instead their info is collected into self.deferred_images so a
            caller can insert them later via WPS UI automation.
            Each entry: {anchor_text, caption, file_path, width, height, position}
            file_path points to an image saved beside output_path in deferred_images/.
            drawing_xml images (template-sourced) set file_path=None and
            include a drawing_xml key instead.

        Returns the absolute path of the generated file.
        """
        self.deferred_images = []
        images_dir = str(Path(output_path).parent / 'deferred_images') if skip_images else None

        tmp_base = tempfile.mkdtemp(prefix='docx_compile_')
        work_dir = os.path.join(tmp_base, 'work')
        try:
            shutil.copytree(self.template_dir, work_dir)
            self._init_rid_counter(work_dir)

            self.abstract_check = self.check_abstract_headings()
            print(f'[compiler] Abstract check: cn={self.abstract_check.has_abstract_cn} en={self.abstract_check.has_abstract_en}')

            print('[compiler] Rebuilding document.xml …')
            stats = self._rebuild_document(work_dir, skip_images=skip_images, images_dir=images_dir)
            print(f'[compiler]   paragraphs : {stats["paragraphs"]}')
            print(f'[compiler]   tables     : {stats["tables"]}')
            print(f'[compiler]   raw_xml    : {stats["raw_xml"]}')
            print(f'[compiler]   images     : {stats["images"]}')
            print(f'[compiler]   deferred   : {stats["deferred"]}')
            print(f'[compiler]   omath      : {stats["omath"]}')
            print(f'[compiler]   ole        : {stats["ole"]}')
            print(f'[compiler]   toc        : {stats["toc"]}')

            print('[compiler] Rebuilding headers/footers …')
            hf_count = self._rebuild_hf_files(work_dir)
            print(f'[compiler]   files rebuilt: {hf_count}')

            self._patch_settings(work_dir)

            _pack_docx(work_dir, output_path)
        finally:
            shutil.rmtree(tmp_base, ignore_errors=True)

        abs_out = os.path.abspath(output_path)
        print(f'[compiler] Done → {abs_out}')
        if skip_images:
            print(f'[compiler] Deferred images : {len(self.deferred_images)}')
        return abs_out

    # ── Abstract heading detection ─────────────────────────────────────────────

    def check_abstract_headings(self) -> AbstractCheckResult:
        """
        Extract all h1 texts (style == "2") and ask the LLM whether
        "摘要" and "Abstract" appear among them.
        """
        h1_texts = [
            e.get('text', '')
            for e in self.ext.get('body_elements', [])
            if e.get('type') == 'paragraph' and e.get('style') == '2'
        ]
        if not h1_texts:
            return AbstractCheckResult(has_abstract_cn=False, has_abstract_en=False)

        heading_list = '\n'.join(f'- {t}' for t in h1_texts)
        return ba.call_structured(
            system_prompt=(
                '你是一个严格的字符串匹配工具，只做精确匹配，不做任何语义推断或联想。'
            ),
            user_prompt=(
                f'以下是文档的全部一级标题（每行一条）：\n{heading_list}\n\n'
                '判断规则（必须严格遵守）：\n'
                '- has_abstract_cn：将某条标题去掉所有空格（含全角空格）后，'
                '结果恰好等于"摘要"，才为 true。否则为 false。\n'
                '- has_abstract_en：将某条标题去掉所有空格后，忽略大小写，'
                '结果恰好等于"abstract"，才为 true（即 Abstract / ABSTRACT / abstract 均算匹配）。否则为 false。\n'
                '不得根据语义、上下文或相似词进行联想。列表中找不到精确匹配就返回 false。'
            ),
            response_model=AbstractCheckResult,
        )

    # ── rId management ─────────────────────────────────────────────────────────

    def _init_rid_counter(self, work_dir: str) -> None:
        rels_path = os.path.join(work_dir, 'word', '_rels', 'document.xml.rels')
        if not os.path.exists(rels_path):
            self._rid_counter = 100
            return
        root = ET.parse(rels_path).getroot()
        nums = [
            int(rel.get('Id', 'rId0')[3:])
            for rel in root
            if rel.get('Id', '').startswith('rId')
            and rel.get('Id', '')[3:].isdigit()
        ]
        self._rid_counter = max(nums, default=0) + 1

    def _alloc_rid(self) -> str:
        rid = f'rId{self._rid_counter}'
        self._rid_counter += 1
        return rid

    def _alloc_shape_id(self) -> int:
        sid = self._shape_counter
        self._shape_counter += 1
        return sid

    # ── Document body reconstruction ───────────────────────────────────────────

    def _rebuild_document(
        self,
        work_dir:   str,
        skip_images: bool = True,
        images_dir:  str | None = None,
    ) -> dict:
        """
        Reconstruct word/document.xml entirely from extraction body_elements.

        The template document.xml contributes:
          • The root <w:document> element with all namespace declarations.
          • The structural XML of each table (tblPr, tblGrid, trPr, tcPr).
          • The body-level <w:sectPr> (appended last, unchanged).

        skip_images / images_dir:
            When skip_images=True, image elements are collected into
            self.deferred_images instead of being written to the XML.
            Base64 images are saved as files under images_dir.
        """
        doc_path = os.path.join(work_dir, 'word', 'document.xml')

        # Parse template to obtain the root element (namespace declarations)
        # and the template body for structural reference.
        tree = ET.parse(doc_path)
        root = tree.getroot()
        body = root.find(_q('body'))

        # Collect template tables (positional: k-th extraction table → k-th
        # template table) and the final body-level sectPr before clearing.
        tmpl_tables = [c for c in body if _tag(c) == 'tbl']
        final_sectPr = body.find(_q('sectPr'))  # body-level, not inside pPr

        # Clear the body — every element will come from extraction.
        for child in list(body):
            body.remove(child)

        stats = dict(paragraphs=0, tables=0, images=0, deferred=0, omath=0, ole=0, toc=0, raw_xml=0)
        table_seq = 0
        img_file_seq = 0

        body_elements_list = self.ext.get('body_elements', [])
        all_para_texts = [
            (e.get('text') or '').strip()
            for e in body_elements_list
            if e.get('type') == 'paragraph'
        ]
        # When deferring images, suppress caption paragraphs that duplicate the
        # image element's own caption field (inserted by docx_tools.insert_figure).
        captions_to_suppress: set[str] = {
            (e.get('caption') or '').strip()
            for e in body_elements_list
            if e.get('type') == 'image' and (e.get('caption') or '').strip()
        } if skip_images else set()

        for elem_idx, elem in enumerate(body_elements_list):
            etype = elem.get('type')

            if etype == 'paragraph':
                para_text = (elem.get('text') or '').strip()
                if para_text and para_text in captions_to_suppress:
                    continue
                body.append(self._build_para(elem))
                stats['paragraphs'] += 1

            elif etype == 'raw_xml':
                # TOC field blocks, fldChar structures, w:sdt — preserved verbatim
                # by the extractor.  Re-emit directly; never attempt to parse or
                # regenerate these — doing so breaks TOC hyperlinks, page numbers,
                # cross-references, and structured-document-tag numbering.
                node = _parse_xml(elem.get('xml', ''))
                if node is not None:
                    body.append(node)
                stats['raw_xml'] += 1

            elif etype == 'toc':
                if self.abstract_check.has_abstract_cn or self.abstract_check.has_abstract_en:
                    body.append(self._build_toc_para(elem))
                    stats['toc'] += 1
                else:
                    print('[compiler] TOC skipped: no abstract heading detected')

            elif etype == 'table':
                tmpl_tbl = (
                    tmpl_tables[table_seq]
                    if table_seq < len(tmpl_tables)
                    else None
                )
                body.append(self._build_table(elem, tmpl_tbl))
                table_seq += 1
                stats['tables'] += 1

            elif etype == 'image':
                if skip_images:
                    _after   = _collect_after_texts(body_elements_list, elem_idx)
                    _caption = _generate_caption_llm(_after, elem.get('caption', ''))
                    _anchor  = _generate_anchor_llm(
                        _collect_before_texts(body_elements_list, elem_idx),
                        all_para_texts,
                    )
                    record: dict = {
                        'anchor_text':  _anchor,
                        'caption':      _caption,
                        '_after_texts': _after,
                        'width':        float(elem.get('width',  0) or 0),
                        'height':       float(elem.get('height', 0) or 0),
                        'position':     elem.get('position', 'center'),
                        'file_path':    None,
                        'drawing_xml':  None,
                    }
                    drawing_xml = elem.get('drawing_xml')
                    raw_b64     = elem.get('base64', '')
                    if drawing_xml:
                        # 从 work_dir/word/media/ 提取真实图片文件
                        copied = self._copy_drawing_image(
                            drawing_xml, work_dir, images_dir, img_file_seq
                        )
                        if copied:
                            record['file_path'] = copied
                            img_file_seq += 1
                        else:
                            record['drawing_xml'] = drawing_xml
                    elif raw_b64:
                        try:
                            img_bytes = base64.b64decode(raw_b64, validate=True)
                            suffix    = _image_suffix(img_bytes)
                            if images_dir:
                                os.makedirs(images_dir, exist_ok=True)
                            save_dir  = images_dir or tempfile.gettempdir()
                            file_path = os.path.join(save_dir, f'img_{img_file_seq:04d}.{suffix}')
                            with open(file_path, 'wb') as fh:
                                fh.write(img_bytes)
                            record['file_path'] = os.path.abspath(file_path)
                            img_file_seq += 1
                        except Exception as exc:
                            print(f'[compiler] deferred image save failed: {exc}')
                    self.deferred_images.append(record)
                    stats['deferred'] += 1
                else:
                    nodes = self._build_image_nodes(elem, work_dir)
                    for n in nodes:
                        body.append(n)
                    stats['images'] += 1

            elif etype in ('omath', 'omathpara'):
                body.append(self._build_omath_para(elem))
                stats['omath'] += 1

            elif etype == 'ole':
                body.append(self._build_ole_para(elem, work_dir))
                stats['ole'] += 1

            # Other tags (bookmarkEnd, etc.) are internal Word markers with no
            # content stored in extraction — skip them safely.

        if skip_images and self.deferred_images:
            _deduplicate_captions(self.deferred_images)
            for _rec in self.deferred_images:
                _rec.pop('_after_texts', None)

        # The body-level sectPr is not fully stored in extraction.json
        # (only header_refs, footer_refs, page_size are captured — not the
        # full XML with margins, line numbers, header distance, etc.).
        # This is a legitimate gap: use the template's sectPr verbatim.
        if final_sectPr is not None:
            final_copy = copy.deepcopy(final_sectPr)
            # Strip <w:type> from the template's body-level sectPr — the template
            # uses oddPage here, which forces a blank even page before the last section.
            for t in final_copy.findall(_q('type')):
                final_copy.remove(t)
            body.append(final_copy)

        _write_xml(doc_path, root)
        return stats

    # ── Section break builder ──────────────────────────────────────────────────

    def _build_sectPr(self, sb: dict) -> ET.Element:
        """
        Build a <w:sectPr> element from a section_break dict.

        sb keys:
            header_refs        : {"default": "rIdX", "even": "rIdY", "first": "rIdZ"}
            footer_refs        : {"default": "rIdX", "even": "rIdY"}
            page_size          : {"w": "11906", "h": "16838"}
            restart_page_number: int | None

        The rIds are written verbatim — they reference header/footer XML files
        already present in the template's word/ directory, so no .rels surgery
        is needed for template-owned headers/footers.

        pgMar values are fixed to the HIT template standard (A4, matching every
        section in template.docx).  <w:type> is intentionally omitted so Word/WPS
        treats the break as "nextPage" (the OOXML default), which forces a page
        break at the paragraph boundary.
        """
        sectPr = ET.Element(_q('sectPr'))

        # Header references
        if sb.get('header_refs', {}):
            for ref_type, rid in sb.get('header_refs', {}).items():
                ref = ET.SubElement(sectPr, _q('headerReference'))
                ref.set(_q('type'), ref_type)
                ref.set(_qr('id'), rid)

        # Footer references
        if sb.get('footer_refs', {}):
            for ref_type, rid in sb.get('footer_refs', {}).items():
                ref = ET.SubElement(sectPr, _q('footerReference'))
                ref.set(_q('type'), ref_type)
                ref.set(_qr('id'), rid)

        # Page size
        ps = sb.get('page_size', {})
        pgSz = ET.SubElement(sectPr, _q('pgSz'))
        pgSz.set(_q('w'), str(ps.get('w', '11906')))
        pgSz.set(_q('h'), str(ps.get('h', '16838')))

        # Page margins — HIT template fixed values
        pgMar = ET.SubElement(sectPr, _q('pgMar'))
        pgMar.set(_q('top'),    '2155')
        pgMar.set(_q('right'),  '1701')
        pgMar.set(_q('bottom'), '1701')
        pgMar.set(_q('left'),   '1701')
        pgMar.set(_q('header'), '1701')
        pgMar.set(_q('footer'), '1304')
        pgMar.set(_q('gutter'), '0')

        # Page number restart
        rp = sb.get('restart_page_number')
        if rp is not None:
            pgNum = ET.SubElement(sectPr, _q('pgNumType'))
            pgNum.set(_q('start'), str(rp))

        # Single column
        cols = ET.SubElement(sectPr, _q('cols'))
        cols.set(_q('space'), '720')

        # Document grid
        docGrid = ET.SubElement(sectPr, _q('docGrid'))
        docGrid.set(_q('type'),      'linesAndChars')
        docGrid.set(_q('linePitch'), '395')
        docGrid.set(_q('charSpace'), '1861')

        return sectPr

    # ── Paragraph builder ──────────────────────────────────────────────────────

    def _build_para(self, pdata: dict) -> ET.Element:
        """
        Build a <w:p> element entirely from extraction paragraph data.

        Source of every piece:
          pPr  → pdata['pPr']  (XML string, re-parsed; ns0: → w: on output)
          rPr  → run['rPr']    (XML string, re-parsed)
          text → run['text']

        If the element carries a 'section_break' dict, a <w:sectPr> is built
        and appended to <w:pPr>.  Omitting <w:type> inside sectPr lets Word/WPS
        default to "nextPage", so the break both changes headers/footers and
        starts a new page.
        """
        p = ET.Element(_q('p'))

        pPr_elem = _parse_xml(pdata.get('pPr'))

        # Attach sectPr when this paragraph marks a section boundary
        sb = pdata.get('section_break')
        if sb is not None:
            if pPr_elem is None:
                pPr_elem = ET.Element(_q('pPr'))
            pPr_elem.append(self._build_sectPr(sb))

        if pPr_elem is not None:
            p.append(pPr_elem)

        for run in pdata.get('runs', []):
            r = ET.SubElement(p, _q('r'))

            rPr_elem = _parse_xml(run.get('rPr'))
            if rPr_elem is not None:
                r.append(rPr_elem)

            if run.get('drawing_xml'):
                drawing_elem = _parse_xml(run['drawing_xml'])
                if drawing_elem is not None:
                    # mc:AlternateContent — stored whole by the updated extractor;
                    # re-emit directly (includes VML fallback for WPS).
                    # Bare w:drawing with a WPS anchor shape — extractor stripped
                    # the mc:AlternateContent wrapper; restore it so WPS honours
                    # the posOffset values instead of floating the shape randomly.
                    if (
                        _tag(drawing_elem) != 'AlternateContent'
                        and _is_wps_anchor_drawing(drawing_elem)
                    ):
                        r.append(_wrap_in_mc_choice(drawing_elem))
                    else:
                        r.append(drawing_elem)
            elif run.get('object_xml'):
                # OLE object (e.g. Equation Editor formula) — re-emit verbatim.
                object_elem = _parse_xml(run['object_xml'])
                if object_elem is not None:
                    r.append(object_elem)
            else:
                text = run.get('text') or ''
                segments = text.split('\n')
                for seg_i, seg in enumerate(segments):
                    t = ET.SubElement(r, _q('t'))
                    t.text = seg
                    if seg and (seg[0] == ' ' or seg[-1] == ' '):
                        t.set(XML_SPACE, 'preserve')
                    if seg_i < len(segments) - 1:
                        ET.SubElement(r, _q('br'))

        return p

    # ── TOC field builder ──────────────────────────────────────────────────────

    def _build_toc_para(self, elem: dict) -> ET.Element:
        """
        Build a single <w:p> that contains a TOC complex field.

        The field uses ``w:dirty="true"`` so Word/WPS regenerates the entire
        Table of Contents from the document's heading paragraphs the first time
        the document is opened.

        The ``\\h`` hyperlink switch is intentionally omitted: with ``\\h``
        Word wraps each entry in a ``<w:hyperlink>`` element and applies the
        Hyperlink character style (blue + underlined).  Without ``\\h`` the
        auto-generated entries are plain black text, which is the standard
        appearance for Chinese academic thesis TOCs.

        Field instruction: ``TOC \\o "1-N" \\z``
            \\o "1-N"  — include heading outline levels 1 through max_level
            \\z        — suppress tab leader and page numbers in Web Layout view
        """
        max_level = elem.get('max_level', 4)
        instr = f' TOC \\o "1-{max_level}" \\z '

        p = ET.Element(_q('p'))

        # begin — w:dirty="true" tells Word the cached content is stale
        r1  = ET.SubElement(p, _q('r'))
        fc1 = ET.SubElement(r1, _q('fldChar'))
        fc1.set(_q('fldCharType'), 'begin')
        fc1.set(_q('dirty'), 'true')

        # field instruction
        r2 = ET.SubElement(p, _q('r'))
        it = ET.SubElement(r2, _q('instrText'))
        it.set(XML_SPACE, 'preserve')
        it.text = instr

        # separate — cached content follows (empty; Word will fill it in)
        r3  = ET.SubElement(p, _q('r'))
        fc3 = ET.SubElement(r3, _q('fldChar'))
        fc3.set(_q('fldCharType'), 'separate')

        # end
        r4  = ET.SubElement(p, _q('r'))
        fc4 = ET.SubElement(r4, _q('fldChar'))
        fc4.set(_q('fldCharType'), 'end')

        return p

    # ── Table builder ──────────────────────────────────────────────────────────

    def _build_table(
        self,
        tdata:    dict,
        tmpl_tbl: ET.Element | None,
    ) -> ET.Element:
        """
        Reconstruct a <w:tbl> element.

        Cell content (paragraphs, text, run formatting) → extraction.json.
        Table skeleton (tblPr, tblGrid, trPr, tcPr)     → template table.

        If no matching template table exists (e.g. a table added by
        DocxTools that was not in the original document), a minimal valid
        skeleton is used instead.
        """
        # Deep-copy the template table so we start with the correct skeleton
        # (tblPr, tblGrid) and row/cell structure (trPr, tcPr).
        if tmpl_tbl is not None:
            tbl = copy.deepcopy(tmpl_tbl)
        else:
            tbl = _minimal_tbl()

        # Detach all rows from the skeleton — we will rebuild them.
        tmpl_rows = tbl.findall(_q('tr'))
        for tr in tmpl_rows:
            tbl.remove(tr)

        ext_rows = tdata.get('rows', [])

        for row_i, ext_row in enumerate(ext_rows):
            # Choose the template row that best represents this row's structure.
            if row_i < len(tmpl_rows):
                row_proto = tmpl_rows[row_i]
            elif tmpl_rows:
                row_proto = tmpl_rows[-1]   # clone last row for extra rows
            else:
                row_proto = ET.Element(_q('tr'))

            tr = copy.deepcopy(row_proto)

            # Detach cells from the cloned row.
            tmpl_cells = tr.findall(_q('tc'))
            for tc in tmpl_cells:
                tr.remove(tc)

            for col_i, ext_cell in enumerate(ext_row):
                # Choose the template cell that best represents this column.
                if col_i < len(tmpl_cells):
                    cell_proto = tmpl_cells[col_i]
                elif tmpl_cells:
                    cell_proto = tmpl_cells[-1]
                else:
                    cell_proto = _minimal_tc()

                tc = copy.deepcopy(cell_proto)

                # Remove existing paragraphs from the cloned cell
                # (leave tcPr intact — it carries column width, borders, etc.).
                for old_p in tc.findall(_q('p')):
                    tc.remove(old_p)

                # Append paragraphs from extraction.
                cell_paras = ext_cell.get('paragraphs', [])
                if cell_paras:
                    for pdata in cell_paras:
                        tc.append(self._build_para(pdata))
                else:
                    # OOXML requires at least one <w:p> per cell.
                    tc.append(ET.Element(_q('p')))

                tr.append(tc)

            tbl.append(tr)

        _rebalance_col_widths(tbl, ext_rows)
        return tbl

    # ── Image paragraph builder ────────────────────────────────────────────────

    def _build_image_nodes(
        self,
        elem:     dict,
        work_dir: str,
    ) -> list[ET.Element]:
        """
        Return a list of <w:p> elements for an image element.

        Two source paths are supported:

        drawing_xml path (template-derived images)
            The extractor stored the original <w:drawing> XML verbatim.
            Re-emit it directly so that all DrawingML properties (positioning,
            size, effects) and the rId → media file mapping are preserved
            exactly.  The rId is valid because the compiler copies the
            template's word/media/ and word/_rels/ files unchanged.

        base64 path (DocxTools-added images)
            Build a fresh inline DrawingML paragraph from the image bytes.
            Falls back to a plain-text placeholder when decoding fails.
        """
        caption  = elem.get('caption', '')
        position = elem.get('position', 'center')
        nodes: list[ET.Element] = []

        # ── drawing_xml path ──────────────────────────────────────────────────
        drawing_xml = elem.get('drawing_xml')
        if drawing_xml:
            drawing_elem = _parse_xml(drawing_xml)
            if drawing_elem is not None:
                p = ET.Element(_q('p'))
                r = ET.SubElement(p, _q('r'))
                r.append(drawing_elem)
                nodes.append(p)
                if caption:
                    nodes.append(self._caption_para(caption))
                return nodes

        # ── base64 path ───────────────────────────────────────────────────────
        raw_b64   = elem.get('base64', '')
        width_pt  = int(elem.get('width',  0) or 0)
        height_pt = int(elem.get('height', 0) or 0)

        img_bytes: bytes | None = None
        try:
            img_bytes = base64.b64decode(raw_b64, validate=True)
        except Exception:
            pass

        if img_bytes is None:
            label = f'[图片: {caption}]' if caption else '[图片]'
            nodes.append(_plain_para(label))
            return nodes

        suffix   = _image_suffix(img_bytes)
        rid      = self._embed_image(img_bytes, suffix, work_dir)
        cx       = (width_pt  or 100) * 12700   # points → EMU
        cy       = (height_pt or 100) * 12700
        shape_id = self._alloc_shape_id()
        img_name = caption or f'image{shape_id}'

        nodes.append(self._drawing_para(rid, cx, cy, shape_id, img_name, position))

        if caption:
            nodes.append(self._caption_para(caption))

        return nodes

    def _drawing_para(
        self,
        rid:      str,
        cx:       int,
        cy:       int,
        shape_id: int,
        name:     str,
        position: str,
    ) -> ET.Element:
        """Build a <w:p> containing an inline <w:drawing> for the given rId."""
        p = ET.Element(_q('p'))

        if position == 'center':
            pPr = ET.SubElement(p, _q('pPr'))
            jc  = ET.SubElement(pPr, _q('jc'))
            jc.set(_q('val'), 'center')

        r       = ET.SubElement(p,       _q('r'))
        drawing = ET.SubElement(r,       _q('drawing'))
        inline  = ET.SubElement(drawing, f'{{{WP}}}inline')
        for attr in ('distT', 'distB', 'distL', 'distR'):
            inline.set(attr, '0')

        extent = ET.SubElement(inline, f'{{{WP}}}extent')
        extent.set('cx', str(cx))
        extent.set('cy', str(cy))

        docPr = ET.SubElement(inline, f'{{{WP}}}docPr')
        docPr.set('id',   str(shape_id))
        docPr.set('name', name)

        ET.SubElement(inline, f'{{{WP}}}cNvGraphicFramePr')

        graphic     = ET.SubElement(inline,      f'{{{A}}}graphic')
        graphicData = ET.SubElement(graphic,     f'{{{A}}}graphicData')
        graphicData.set('uri', PIC)

        pic_e    = ET.SubElement(graphicData, f'{{{PIC}}}pic')
        nvPicPr  = ET.SubElement(pic_e,       f'{{{PIC}}}nvPicPr')
        cNvPr    = ET.SubElement(nvPicPr,     f'{{{PIC}}}cNvPr')
        cNvPr.set('id',   '0')
        cNvPr.set('name', name)
        ET.SubElement(nvPicPr, f'{{{PIC}}}cNvPicPr')

        blipFill = ET.SubElement(pic_e,    f'{{{PIC}}}blipFill')
        blip     = ET.SubElement(blipFill, f'{{{A}}}blip')
        blip.set(f'{{{R}}}embed', rid)
        stretch  = ET.SubElement(blipFill, f'{{{A}}}stretch')
        ET.SubElement(stretch, f'{{{A}}}fillRect')

        spPr    = ET.SubElement(pic_e, f'{{{PIC}}}spPr')
        xfrm    = ET.SubElement(spPr,  f'{{{A}}}xfrm')
        off     = ET.SubElement(xfrm,  f'{{{A}}}off')
        off.set('x', '0');  off.set('y', '0')
        ext_e   = ET.SubElement(xfrm,  f'{{{A}}}ext')
        ext_e.set('cx', str(cx));  ext_e.set('cy', str(cy))
        prstGeom = ET.SubElement(spPr, f'{{{A}}}prstGeom')
        prstGeom.set('prst', 'rect')
        ET.SubElement(prstGeom, f'{{{A}}}avLst')

        return p

    def _caption_para(self, caption: str) -> ET.Element:
        """Plain centered paragraph used as a figure/table caption."""
        p   = ET.Element(_q('p'))
        pPr = ET.SubElement(p, _q('pPr'))
        jc  = ET.SubElement(pPr, _q('jc'))
        jc.set(_q('val'), 'center')
        r = ET.SubElement(p, _q('r'))
        t = ET.SubElement(r, _q('t'))
        t.text = caption
        return p

    def _embed_image(self, img_bytes: bytes, suffix: str, work_dir: str) -> str:
        """
        Write image bytes to word/media/, register the relationship, update
        [Content_Types].xml, and return the allocated rId.
        """
        media_dir = os.path.join(work_dir, 'word', 'media')
        os.makedirs(media_dir, exist_ok=True)

        existing = os.listdir(media_dir)
        nums = [
            int(m.group(1))
            for f in existing
            if (m := re.match(r'image(\d+)', f))
        ]
        next_num  = max(nums, default=0) + 1
        img_name  = f'image{next_num}{suffix}'
        with open(os.path.join(media_dir, img_name), 'wb') as f:
            f.write(img_bytes)

        rid       = self._alloc_rid()
        rels_path = os.path.join(work_dir, 'word', '_rels', 'document.xml.rels')
        _rels_append(rels_path, rid, REL_TYPE_IMAGE, f'media/{img_name}')

        ct_path = os.path.join(work_dir, '[Content_Types].xml')
        _ensure_content_type(ct_path, suffix)

        return rid

    # ── OLE paragraph builder ──────────────────────────────────────────────────

    def _build_ole_para(self, elem: dict, work_dir: str) -> ET.Element:
        """
        Build a paragraph containing an OLE embedded object (e.g. Equation.3).

        Layout:
          - Inline (text_before or text_after present):
              text_before run + OLE object run + text_after run, all in one paragraph.
          - Block with formula_index: tab-stop layout — center tab (4252 twips) pushes
            the equation to the middle; right tab (8504 twips) pushes the index
            number to the right margin.
          - Block without formula_index: simple jc="center" paragraph.
        """
        raw_b64       = elem.get('base64', '')
        formula_index = (elem.get('formula_index') or '').strip()
        width_pt      = float(elem.get('width_pt',  120))
        height_pt     = float(elem.get('height_pt',  30))
        prog_id       = elem.get('prog_id', 'Equation.3')
        text_before   = (elem.get('text_before') or '').strip()
        text_after    = (elem.get('text_after')  or '').strip()
        is_inline     = bool(text_before or text_after or elem.get('is_inline'))

        try:
            ole_bytes = base64.b64decode(raw_b64, validate=True)
        except Exception:
            fallback_label = formula_index or '[OLE 对象]'
            fallback_text  = f'{text_before}{fallback_label}{text_after}'.strip() or fallback_label
            return _plain_para(fallback_text)

        ole_rid    = self._embed_ole(ole_bytes, work_dir)
        shape_id   = self._alloc_shape_id()
        shape_name = f'_x0000_i{shape_id + 1024}'

        p   = ET.Element(_q('p'))
        pPr = ET.SubElement(p, _q('pPr'))

        if is_inline:
            # Inline: no special paragraph alignment; text runs surround the object.
            if text_before:
                r_before = ET.SubElement(p, _q('r'))
                t_before = ET.SubElement(r_before, _q('t'))
                t_before.text = text_before
                t_before.set(XML_SPACE, 'preserve')
        elif formula_index:
            # Block with index: tab-stop layout.
            # HIT template: content width = 11906 - 1701×2 = 8504 twips.
            tabs_elem = ET.SubElement(pPr, _q('tabs'))
            tab_c = ET.SubElement(tabs_elem, _q('tab'))
            tab_c.set(_q('val'), 'center')
            tab_c.set(_q('pos'), '4252')
            tab_r = ET.SubElement(tabs_elem, _q('tab'))
            tab_r.set(_q('val'), 'right')
            tab_r.set(_q('pos'), '8504')
            r_tab1 = ET.SubElement(p, _q('r'))
            ET.SubElement(r_tab1, _q('tab'))
        else:
            jc = ET.SubElement(pPr, _q('jc'))
            jc.set(_q('val'), 'center')

        # OLE object run
        r   = ET.SubElement(p, _q('r'))
        obj = ET.SubElement(r, _q('object'))
        obj.set(_q('dxaOrig'), str(int(width_pt  * 20)))
        obj.set(_q('dyaOrig'), str(int(height_pt * 20)))

        shape = ET.SubElement(obj, f'{{{V_NS}}}shape')
        shape.set('id',      shape_name)
        shape.set('type',    '#_x0000_t75')
        shape.set('style',   f'width:{width_pt}pt;height:{height_pt}pt')
        shape.set(f'{{{O_NS}}}ole', '')
        shape.set('stroked', 'f')

        img_b64 = elem.get('image_base64', '')
        if img_b64:
            try:
                img_bytes = base64.b64decode(img_b64, validate=True)
                img_rid   = self._embed_image(img_bytes, 'wmf', work_dir)
                imgdata   = ET.SubElement(shape, f'{{{V_NS}}}imagedata')
                imgdata.set(f'{{{R}}}id',       img_rid)
                imgdata.set(f'{{{O_NS}}}title', '')
            except Exception:
                pass

        ole_elem = ET.SubElement(obj, f'{{{O_NS}}}OLEObject')
        ole_elem.set('Type',       'Embed')
        ole_elem.set('ProgID',     prog_id)
        ole_elem.set('ShapeID',    shape_name)
        ole_elem.set('DrawAspect', 'Content')
        ole_elem.set('ObjectID',   f'_{shape_id}')
        ole_elem.set(f'{{{R}}}id', ole_rid)

        if is_inline:
            if text_after:
                r_after = ET.SubElement(p, _q('r'))
                t_after = ET.SubElement(r_after, _q('t'))
                t_after.text = text_after
                t_after.set(XML_SPACE, 'preserve')
        else:
            if formula_index:
                r_tab2 = ET.SubElement(p, _q('r'))
                ET.SubElement(r_tab2, _q('tab'))
                r_idx = ET.SubElement(p, _q('r'))
                t_idx = ET.SubElement(r_idx, _q('t'))
                t_idx.text = formula_index

        return p

    def _copy_drawing_image(
        self,
        drawing_xml: str,
        work_dir:    str,
        images_dir:  str | None,
        seq:         int,
    ) -> str | None:
        """
        Extract the image file referenced by drawing_xml from work_dir/word/media/
        and copy it to images_dir.  Returns the absolute destination path, or None
        on failure.
        """
        try:
            # Find r:embed rId in the drawing XML
            rid_match = re.search(r'r:embed="(rId\d+)"', drawing_xml)
            if not rid_match:
                return None
            rid = rid_match.group(1)

            # Look up the rId in document.xml.rels
            rels_path = os.path.join(work_dir, 'word', '_rels', 'document.xml.rels')
            rels_tree = ET.parse(rels_path)
            target = None
            for rel in rels_tree.getroot():
                if rel.get('Id') == rid:
                    target = rel.get('Target')  # e.g. "media/image1.tiff"
                    break
            if not target:
                return None

            src = os.path.join(work_dir, 'word', target)
            if not os.path.exists(src):
                return None

            ext      = os.path.splitext(target)[1].lstrip('.') or 'png'
            save_dir = images_dir or tempfile.gettempdir()
            os.makedirs(save_dir, exist_ok=True)
            dst = os.path.join(save_dir, f'img_{seq:04d}.{ext}')
            shutil.copy2(src, dst)
            return os.path.abspath(dst)
        except Exception as exc:
            print(f'[compiler] _copy_drawing_image failed: {exc}')
            return None

    def _embed_ole(self, ole_bytes: bytes, work_dir: str) -> str:
        """Write OLE bytes to word/embeddings/ and register the relationship."""
        emb_dir = os.path.join(work_dir, 'word', 'embeddings')
        os.makedirs(emb_dir, exist_ok=True)

        existing = os.listdir(emb_dir)
        nums = [
            int(m.group(1))
            for f in existing
            if (m := re.match(r'oleObject(\d+)\.bin', f))
        ]
        next_num = max(nums, default=0) + 1
        ole_name = f'oleObject{next_num}.bin'
        with open(os.path.join(emb_dir, ole_name), 'wb') as f:
            f.write(ole_bytes)

        rid       = self._alloc_rid()
        rels_path = os.path.join(work_dir, 'word', '_rels', 'document.xml.rels')
        _rels_append(rels_path, rid, REL_TYPE_OLE, f'embeddings/{ole_name}')
        return rid

    # ── OMath paragraph builder ────────────────────────────────────────────────

    def _build_omath_para(self, elem: dict) -> ET.Element:
        """
        Build a <w:p> containing an Office Math (OMML) block or inline formula.

        Inline (text_before or text_after present):
            text_before run + <m:oMath> + text_after run in one paragraph.
        Block:
            <m:oMathPara><m:oMath>…</m:oMathPara> with optional formula_index run.
        """
        formula       = (elem.get('formula') or '').strip()
        formula_index = (elem.get('formula_index') or '').strip()
        text_before   = (elem.get('text_before') or '').strip()
        text_after    = (elem.get('text_after')  or '').strip()
        is_inline     = bool(text_before or text_after or elem.get('is_inline'))

        p = ET.Element(_q('p'))

        if is_inline:
            if text_before:
                r_before = ET.SubElement(p, _q('r'))
                t_before = ET.SubElement(r_before, _q('t'))
                t_before.text = text_before
                t_before.set(XML_SPACE, 'preserve')

            if formula.startswith('<'):
                formula_elem = _parse_xml(formula)
                if formula_elem is not None:
                    if _tag(formula_elem) == 'oMath':
                        p.append(formula_elem)
                    else:
                        oMath = ET.SubElement(p, _qm('oMath'))
                        oMath.append(formula_elem)
                else:
                    oMath = ET.SubElement(p, _qm('oMath'))
                    _math_text(oMath, formula)
            else:
                oMath = ET.SubElement(p, _qm('oMath'))
                _math_text(oMath, formula)

            if text_after:
                r_after = ET.SubElement(p, _q('r'))
                t_after = ET.SubElement(r_after, _q('t'))
                t_after.text = text_after
                t_after.set(XML_SPACE, 'preserve')
        else:
            if formula.startswith('<'):
                formula_elem = _parse_xml(formula)
                if formula_elem is not None:
                    ftag = _tag(formula_elem)
                    if ftag == 'oMathPara':
                        p.append(formula_elem)
                    elif ftag == 'oMath':
                        oMathPara = ET.SubElement(p, _qm('oMathPara'))
                        oMathPara.append(formula_elem)
                    else:
                        oMathPara = ET.SubElement(p, _qm('oMathPara'))
                        oMath = ET.SubElement(oMathPara, _qm('oMath'))
                        oMath.append(formula_elem)
                else:
                    oMathPara = ET.SubElement(p, _qm('oMathPara'))
                    oMath = ET.SubElement(oMathPara, _qm('oMath'))
                    _math_text(oMath, formula)
            else:
                oMathPara = ET.SubElement(p, _qm('oMathPara'))
                oMath = ET.SubElement(oMathPara, _qm('oMath'))
                _math_text(oMath, formula)

            if formula_index:
                r = ET.SubElement(p, _q('r'))
                t = ET.SubElement(r, _q('t'))
                t.text = f'  {formula_index}'
                t.set(XML_SPACE, 'preserve')

        return p

    # ── settings.xml patch ─────────────────────────────────────────────────────

    def _patch_settings(self, work_dir: str) -> None:
        """
        Ensure ``word/settings.xml`` contains ``<w:updateFields w:val="true"/>``.

        This instructs Word/WPS to recalculate all fields (including the TOC)
        every time the document is opened, so the TOC is always up to date
        without any manual intervention from the user.

        Uses string-based injection to avoid disturbing existing namespace
        declarations in the settings file.
        """
        settings_path = os.path.join(work_dir, 'word', 'settings.xml')
        if not os.path.exists(settings_path):
            return
        with open(settings_path, 'r', encoding='utf-8') as f:
            text = f.read()
        if 'updateFields' in text:
            return
        entry = '<w:updateFields w:val="true"/>'
        text  = text.replace('</w:settings>', entry + '</w:settings>')
        with open(settings_path, 'w', encoding='utf-8') as f:
            f.write(text)

    # ── Header / footer reconstruction ─────────────────────────────────────────

    def _rebuild_hf_files(self, work_dir: str) -> int:
        """
        Rebuild every header and footer XML file from extraction data.

        Unlike the body, headers and footers contain no tables in this
        template, so they can be fully reconstructed: the template root
        element (<w:hdr> / <w:ftr>) is kept for its tag and namespace
        declarations; all children are replaced from extraction.
        """
        word_dir = os.path.join(work_dir, 'word')
        count    = 0

        all_hf: dict[str, dict] = {
            **self.ext.get('headers', {}),
            **self.ext.get('footers', {}),
        }

        for filename, hf_data in all_hf.items():
            xml_path = os.path.join(word_dir, filename)
            if not os.path.exists(xml_path):
                continue
            self._rebuild_hf(xml_path, hf_data)
            count += 1

        return count

    def _rebuild_hf(self, xml_path: str, hf_data: dict) -> None:
        """
        Reconstruct a single header/footer XML file.

        The root element (<w:hdr> or <w:ftr>) comes from the template so
        that its namespace declarations are preserved.  All children are
        replaced with paragraphs built from extraction data.

        Source of content: hf_data['paragraphs'] from extraction.json.
        Source of structure: root element tag only (from template).
        """
        tree = ET.parse(xml_path)
        root = tree.getroot()

        # Remove all existing children — we rebuild entirely from extraction.
        for child in list(root):
            root.remove(child)

        for pdata in hf_data.get('paragraphs', []):
            root.append(self._build_para(pdata))

        # OOXML requires at least one paragraph in a header/footer.
        if not hf_data.get('paragraphs'):
            root.append(ET.Element(_q('p')))

        _write_xml(xml_path, root)


# ── Module-level structural helpers ───────────────────────────────────────────

def _visual_width(s: str) -> int:
    """CJK characters count as 2 units, everything else as 1."""
    return sum(2 if '一' <= c <= '鿿' else 1 for c in s)


def _max_line_width(cell_text: str) -> int:
    """
    Visual width of the widest 'virtual line' in a cell.

    Two normalisation steps:
    1. Splits on real newlines and literal \\n.
    2. For mixed CJK+ASCII lines (Chinese academic table-header convention:
       CJK label on one virtual line, ASCII/symbol unit on another),
       treats the CJK portion and non-CJK portion as separate virtual lines
       and returns the wider of the two.
    """
    text = (cell_text or '').replace('\\n', '\n')
    lines = text.split('\n')
    max_w = 0
    for line in lines:
        if not line:
            continue
        cjk_w = sum(2 for c in line if '一' <= c <= '鿿')
        rest_w = sum(1 for c in line if not ('一' <= c <= '鿿'))
        if cjk_w > 0 and rest_w > 0:
            # Mixed CJK+ASCII: treat each part as a separate virtual line
            w = max(cjk_w, rest_w)
        else:
            w = cjk_w or rest_w
        if w > max_w:
            max_w = w
    return max_w or 1


def _rebalance_col_widths(tbl: ET.Element, ext_rows: list) -> None:
    """
    Redistribute column widths proportionally by cell content.

    Weight per column = max visual width of any line in any cell of that column.
    Total width is fixed at _TABLE_TEXT_WIDTH_TWIPS.
    Updates both <w:tblGrid> and every <w:tcPr>/<w:tcW> in-place.
    """
    if not ext_rows:
        return
    num_cols = max(len(row) for row in ext_rows)
    if num_cols == 0:
        return

    weights = [0] * num_cols
    for row in ext_rows:
        for col_i, cell in enumerate(row):
            if col_i >= num_cols:
                break
            col_w = _max_line_width(cell.get('text', ''))
            if col_w > weights[col_i]:
                weights[col_i] = col_w

    total_weight = sum(weights) or num_cols
    if total_weight == 0:
        weights = [1] * num_cols
        total_weight = num_cols

    col_widths = [max(1, int(w / total_weight * _TABLE_TEXT_WIDTH_TWIPS)) for w in weights]
    # Absorb rounding error into the last column
    col_widths[-1] += _TABLE_TEXT_WIDTH_TWIPS - sum(col_widths)

    # Rewrite <w:tblGrid>
    tblPr   = tbl.find(_q('tblPr'))
    tblGrid = tbl.find(_q('tblGrid'))
    if tblGrid is None:
        tblGrid = ET.Element(_q('tblGrid'))
        insert_at = (list(tbl).index(tblPr) + 1) if tblPr is not None else 0
        tbl.insert(insert_at, tblGrid)
    else:
        for gc in tblGrid.findall(_q('gridCol')):
            tblGrid.remove(gc)
    for w in col_widths:
        gc = ET.SubElement(tblGrid, _q('gridCol'))
        gc.set(_q('w'), str(w))

    # Rewrite <w:tcPr>/<w:tcW> for every cell
    for tr in tbl.findall(_q('tr')):
        for col_i, tc in enumerate(tr.findall(_q('tc'))):
            if col_i >= num_cols:
                break
            tcPr = tc.find(_q('tcPr'))
            if tcPr is None:
                tcPr = ET.Element(_q('tcPr'))
                tc.insert(0, tcPr)
            tcW = tcPr.find(_q('tcW'))
            if tcW is None:
                tcW = ET.SubElement(tcPr, _q('tcW'))
            tcW.set(_q('w'),    str(col_widths[col_i]))
            tcW.set(_q('type'), 'dxa')


def _minimal_tbl() -> ET.Element:
    """Bare <w:tbl> with the minimum valid structure."""
    tbl    = ET.Element(_q('tbl'))
    tblPr  = ET.SubElement(tbl, _q('tblPr'))
    tblSty = ET.SubElement(tblPr, _q('tblStyle'))
    tblSty.set(_q('val'), 'TableGrid')
    tblW   = ET.SubElement(tblPr, _q('tblW'))
    tblW.set(_q('w'),    '0')
    tblW.set(_q('type'), 'auto')
    ET.SubElement(tbl, _q('tblGrid'))
    return tbl


def _minimal_tc() -> ET.Element:
    """Bare <w:tc> with empty <w:tcPr>."""
    tc = ET.Element(_q('tc'))
    ET.SubElement(tc, _q('tcPr'))
    return tc


def _plain_para(text: str) -> ET.Element:
    """Plain paragraph with a single unstyled run."""
    p = ET.Element(_q('p'))
    r = ET.SubElement(p, _q('r'))
    t = ET.SubElement(r, _q('t'))
    t.text = text
    return p


def _math_text(oMath: ET.Element, text: str) -> None:
    """Append a plain-text run inside an <m:oMath> element."""
    mr = ET.SubElement(oMath, f'{{{M}}}r')
    mt = ET.SubElement(mr,    f'{{{M}}}t')
    mt.text = text


# ── Image meta helpers ─────────────────────────────────────────────────────────

class _CaptionList(BaseModel):
    captions: list[str]


def _collect_before_texts(
    body_elements: list[dict],
    img_idx: int,
    window: int = 5,
) -> list[str]:
    texts: list[str] = []
    for i in range(max(0, img_idx - window), img_idx):
        elem = body_elements[i]
        if elem.get('type') == 'paragraph':
            text = (elem.get('text') or '').strip()
            if text:
                texts.append(text)
    return texts


def _collect_after_texts(
    body_elements: list[dict],
    img_idx: int,
    window: int = 5,
) -> list[str]:
    texts: list[str] = []
    for i in range(img_idx + 1, min(len(body_elements), img_idx + window + 1)):
        elem = body_elements[i]
        if elem.get('type') == 'paragraph':
            text = (elem.get('text') or '').strip()
            if text:
                texts.append(text)
    return texts


def _generate_caption_llm(after_texts: list[str], existing_caption: str) -> str:
    system_prompt = "你是一名学术论文排版助手，擅长为图片添加规范的中文图题。"
    cap_hint = f'已知该图的图题为："{existing_caption}"，' if existing_caption else ''
    if after_texts:
        ctx = '\n'.join(f'[{i + 1}] {t}' for i, t in enumerate(after_texts))
        user_prompt = (
            f"以下是文档中某张图片之后的文本段落：\n{ctx}\n\n"
            f"{cap_hint}请根据上下文为这张图片提供一个规范的中文图题（格式如'图X-X xxx示意图'）。"
            "如上下文中已有明确图题则直接提取，否则根据语义生成。只返回图题文本，不要任何解释。"
        )
    else:
        user_prompt = (
            f"{cap_hint}文档中有一张图片，没有可参考的上下文。"
            "请为其生成一个通用的中文图题（如'示意图'）。只返回图题文本，不要任何解释。"
        )
    try:
        return ba.call(system_prompt, user_prompt).strip() or existing_caption or '示意图'
    except Exception as exc:
        print(f'[compiler] caption LLM call failed: {exc}')
        return existing_caption or '示意图'


# 纯符号/标点文本识别：匹配全由空白、CJK 标点、省略号、ASCII 标点组成的字符串
_SYMBOL_RE = re.compile(
    r'^[\s　'
    r'。，、；：！？'   # 。，、；：！？
    r'…—–·'                      # …—–·
    r'．（）［］'                # ．（）［］
    r'「」『』【】'          # 「」『』【】
    r'‘’“”'                      # ''""
    r'\.,:;!?\-\*\#\(\)\[\]\|/\\@\$\%\^&~`'         # ASCII 标点
    r']+$'
)


def _is_meaningful_text(text: str) -> bool:
    stripped = (text or '').strip()
    return bool(stripped) and not _SYMBOL_RE.match(stripped)


class _AnchorIndex(BaseModel):
    index: int  # 1-based index into the provided list; 0 = none suitable


def _generate_anchor_llm(before_texts: list[str], all_para_texts: list[str]) -> str:
    meaningful = [t for t in before_texts if _is_meaningful_text(t)]
    if not meaningful:
        return ''

    def _unique(t: str) -> bool:
        return sum(1 for p in all_para_texts if t in p) == 1

    system_prompt = "你是一名学术论文排版助手。"
    ctx = '\n'.join(f'[{i + 1}] {t}' for i, t in enumerate(meaningful))
    user_prompt = (
        f"以下是文档中某张图片之前的文本段落：\n{ctx}\n\n"
        "请选出最靠近图片位置、最适合作为插入定位点的段落编号（方括号内的数字）。\n"
        "只返回编号，不要任何解释。如果没有合适的段落，返回 0。"
    )
    try:
        result = ba.call_structured(system_prompt, user_prompt, _AnchorIndex)
        idx = result.index
        if 1 <= idx <= len(meaningful) and _unique(meaningful[idx - 1]):
            return meaningful[idx - 1]
    except Exception as exc:
        print(f'[compiler] anchor LLM call failed: {exc}')

    # Fallback：从最近的有意义段落里找第一个在全文唯一的
    for t in reversed(meaningful):
        if _unique(t):
            return t
    return ''


def _deduplicate_captions(records: list[dict]) -> None:
    from collections import defaultdict
    groups: dict[str, list[int]] = defaultdict(list)
    for i, rec in enumerate(records):
        cap = (rec.get('caption') or '').strip()
        if cap:
            groups[cap].append(i)

    for cap, indices in groups.items():
        if len(indices) <= 1:
            continue
        ctx_parts = []
        for rank, idx in enumerate(indices, 1):
            after_texts = records[idx].get('_after_texts') or []
            ctx = '\n'.join(f'  - {t}' for t in after_texts) or '  （无上下文）'
            ctx_parts.append(f"图片{rank}（当前图题：{cap}）\n后文：\n{ctx}")
        system_prompt = "你是一名学术论文排版助手。"
        user_prompt = (
            f"以下 {len(indices)} 张图片的图题完全相同，请根据各自后文上下文将它们改写为不重复的图题。\n\n"
            + '\n\n'.join(ctx_parts) + '\n\n'
            f"要求：按顺序为每张图片输出一个新图题，格式如'图X-X xxx示意图'，"
            f"共 {len(indices)} 个，每个图题单独一项。"
        )
        try:
            result = ba.call_structured(system_prompt, user_prompt, _CaptionList)
            for rank, idx in enumerate(indices):
                if rank < len(result.captions) and result.captions[rank].strip():
                    records[idx]['caption'] = result.captions[rank].strip()
        except Exception as exc:
            print(f'[compiler] caption dedup LLM call failed: {exc}')


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    _extraction = sys.argv[1] if len(sys.argv) > 1 else 'docx_manager\\docx_engine\\data\\user_extraction.json'
    _output     = sys.argv[2] if len(sys.argv) > 2 else 'docx_manager\\docx_engine\\outputs\\output.docx'
    _template   = sys.argv[3] if len(sys.argv) > 3 else 'docx_manager\\docx_engine\\templates\\hit-template'

    DocxCompiler(
        extraction_path=_extraction,
        template_dir=_template,
    ).compile(output_path=_output)
