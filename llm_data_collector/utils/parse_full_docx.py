from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table
import base64
import io
from typing import List, Dict, Any, Optional
import xml.etree.ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def get_paragraph_text(paragraph: Paragraph) -> str:
    return paragraph.text.strip()


def get_element_text(element) -> str:
    if element is None:
        return ""
    text_parts = []
    for child in element.iter():
        if child.text:
            text_parts.append(child.text)
    return "".join(text_parts).strip()


def get_or_create(element, tag):
    child = element.find(f"{{{W}}}{tag}")
    if child is None:
        child = OxmlElement(f"{{{W}}}{tag}")
        element.append(child)
    return child


def _find_child(element, tag):
    return element.find(f"{{{W}}}{tag}")


def _find_all(element, tag):
    return element.findall(f"{{{W}}}{tag}")


def _find_descendants(element, tag_prefix, tag_name):
    results = []
    for elem in element.iter():
        if elem.tag == f"{{{tag_prefix}}}{tag_name}":
            results.append(elem)
    return results


def parse_heading(paragraph: Paragraph) -> Optional[Dict[str, Any]]:
    style_name = paragraph.style.name if paragraph.style else ""

    if style_name.lower().startswith("heading"):
        try:
            level = int(''.join(filter(str.isdigit, style_name)))
        except (ValueError, AttributeError):
            level = 1

        return {
            "type": f"heading{level}",
            "value": get_paragraph_text(paragraph)
        }

    style_lower = style_name.lower()
    text = get_paragraph_text(paragraph)

    if not text:
        return None

    if style_lower in ("abstract", "摘要", "abstract text", "摘要文本"):
        return {"type": "abstract", "value": text}

    if style_lower in ("conclusion", "结论", "conclusion text", "结论文本"):
        return {"type": "conclusion", "value": text}

    if style_lower in ("acknowledgement", "致谢", "acknowledgement text", "致谢文本"):
        return {"type": "acknowledgement", "value": text}

    if style_lower in ("reference", "参考文献", "references", "参考文献文本"):
        return {"type": "references", "value": text}

    if "标题" in style_name or "title" in style_lower:
        if "摘要" in text or "abstract" in text.lower():
            if "english" in text.lower() or "英文" in text or "en" in text.lower():
                return {"type": "abstract_en", "value": text}
            return {"type": "abstract", "value": text}

        if "结论" in text or "conclusion" in text.lower():
            return {"type": "conclusion", "value": text}

        if "致谢" in text or "acknowledgement" in text.lower():
            return {"type": "acknowledgement", "value": text}

        if "参考文献" in text or "reference" in text.lower():
            return {"type": "references", "value": text}

    return None


def parse_table(table: Table) -> Dict[str, Any]:
    table_data = []
    for row in table.rows:
        row_data = []
        for cell in row.cells:
            cell_text = cell.text.strip()
            row_data.append(cell_text)
        table_data.append(row_data)

    caption = ""
    prev_para = table._element.getprevious()
    if prev_para is not None:
        try:
            prev_p = Paragraph(prev_para, table._tbl._parent)
            caption = prev_p.text.strip()
        except Exception:
            pass

    return {
        "type": "table",
        "caption": caption if caption else None,
        "data": table_data
    }


def _iter_elements_by_tag(element, tag_local_name):
    for child in element.iter():
        if child.tag.endswith("}" + tag_local_name):
            yield child


def parse_image(paragraph: Paragraph, doc: Document) -> Optional[Dict[str, Any]]:
    for run in paragraph.runs:
        for drawing in _iter_elements_by_tag(run._element, "drawing"):
            for blip_fill in _iter_elements_by_tag(drawing, "blipFill"):
                for pic in _iter_elements_by_tag(blip_fill, "pic"):
                    for pic_fill in _iter_elements_by_tag(pic, "picFills"):
                        for blip in _iter_elements_by_tag(pic_fill, "blip"):
                            embed_attr = blip.get(qn("r:embed"))
                            if embed_attr:
                                try:
                                    image_part = doc.part.related_parts.get(embed_attr)
                                    if image_part:
                                        image_bytes = image_part.blob
                                        ext = image_part.content_type.split('/')[-1]
                                        if ext == 'jpeg':
                                            ext = 'jpg'
                                        base64_str = base64.b64encode(image_bytes).decode('utf-8')

                                        caption = get_paragraph_text(paragraph)

                                        width_emus = None
                                        inline_elem = _find_child(drawing, "inline")
                                        if inline_elem is not None:
                                            extent = _find_child(inline_elem, "extent")
                                            if extent is not None:
                                                cx = extent.get(qn("w:cx"))
                                                if cx:
                                                    width_emus = int(cx)
                                        else:
                                            anchor_elem = _find_child(drawing, "anchor")
                                            if anchor_elem is not None:
                                                extent = _find_child(anchor_elem, "extent")
                                                if extent is not None:
                                                    cx = extent.get(qn("w:cx"))
                                                    if cx:
                                                        width_emus = int(cx)

                                        width_inches = None
                                        if width_emus:
                                            width_inches = width_emus / 914400

                                        align = "center"
                                        anchor_elem = _find_child(drawing, "anchor")
                                        if anchor_elem is not None:
                                            position = _find_child(anchor_elem, "position")
                                            if position is not None:
                                                align_elem = _find_child(position, "align")
                                                if align_elem is not None:
                                                    align = align_elem.get(qn("w:val"), "center")

                                        return {
                                            "type": "image",
                                            "base64": base64_str,
                                            "ext": ext,
                                            "caption": caption if caption and not caption.startswith("图") else None,
                                            "width": width_inches,
                                            "align": align
                                        }
                                except Exception:
                                    pass

    return None


def parse_formula(paragraph: Paragraph, doc: Document) -> Optional[Dict[str, Any]]:
    for omath_para in _iter_elements_by_tag(paragraph._element, "oMathPara"):
        omml_str = ET.tostring(omath_para, encoding="unicode", method="xml")

        label = ""
        prev_text = ""
        prev_sibling = paragraph._element.getprevious()
        if prev_sibling is not None:
            try:
                prev_para = Paragraph(prev_sibling, doc)
                prev_text = prev_para.text.strip()
                if "(" in prev_text and ")" in prev_text:
                    label = prev_text
            except Exception:
                pass

        return {
            "type": "formula",
            "omml": omml_str,
            "label": label if label else None
        }

    return None


def parse_references(paragraph: Paragraph, ref_start: bool) -> List[Dict[str, Any]]:
    references = []
    if not ref_start:
        return references

    text = get_paragraph_text(paragraph)
    if not text:
        return references

    import re
    ref_pattern = re.compile(r'\[(\d+)\](.+?)(?=\[|$)', re.DOTALL)
    matches = ref_pattern.findall(text)

    for ref_id, ref_text in matches:
        references.append({
            "type": "reference",
            "id": int(ref_id),
            "text": ref_text.strip()
        })

    return references


def parse_full_docx(doc_path: str) -> Dict[str, Any]:
    doc = Document(doc_path)
    content = []
    headings = []
    ref_started = False

    for element in doc.element.body:
        if element.tag == f"{{{W}}}p":
            paragraph = Paragraph(element, doc)
            heading_result = parse_heading(paragraph)
            if heading_result:
                content.append(heading_result)
                headings.append(heading_result)
                continue

            formula_result = parse_formula(paragraph, doc)
            if formula_result:
                content.append(formula_result)
                continue

            image_result = parse_image(paragraph, doc)
            if image_result:
                content.append(image_result)
                continue

            text = get_paragraph_text(paragraph)
            if "参考文献" in text or "references" in text.lower():
                ref_started = True
                continue

            if ref_started:
                refs = parse_references(paragraph, ref_started)
                content.extend(refs)
                continue

            if text:
                content.append({
                    "type": "body",
                    "value": text
                })

        elif element.tag == f"{{{W}}}tbl":
            table = Table(element, doc.element.body)
            table_result = parse_table(table)
            content.append(table_result)

    return {
        "headings": headings,
        "content": content
    }


def parse_full_docx_simple(doc_path: str) -> List[Dict[str, Any]]:
    result = parse_full_docx(doc_path)
    return result["content"]


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        doc_path = sys.argv[1]
    else:
        doc_path = "data/full_template_v7.docx"

    result = parse_full_docx(doc_path)
    import json
    print(json.dumps(result, ensure_ascii=False, indent=2))
