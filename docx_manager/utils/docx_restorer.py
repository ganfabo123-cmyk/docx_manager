"""
DOCX文档还原器
从JSON格式还原为docx文档
"""
import json
import base64
import io
import traceback
import copy
from typing import Dict, List, Any, Optional
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
V_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class DocxRestorer:
    """DOCX文档还原器"""
    
    HEADING_STYLE_MAP = {
        1: "Heading 1",
        2: "Heading 2", 
        3: "Heading 3"
    }
    
    ALIGN_MAP = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
    }
    
    def __init__(self, elements: List[Dict[str, Any]]):
        self.elements = elements
        self.doc = Document()
        self._ole_counter = 0
        self._shape_counter = 0
    
    def _apply_paragraph_style(self, paragraph, style: Dict[str, Any]):
        style_name = style.get("style_name", "Normal")
        
        if style.get("is_heading") and style.get("heading_level"):
            level = style["heading_level"]
            heading_style = self.HEADING_STYLE_MAP.get(level)
            if heading_style:
                try:
                    paragraph.style = heading_style
                except Exception:
                    pass
        
        alignment = style.get("alignment")
        if alignment and alignment in self.ALIGN_MAP:
            paragraph.alignment = self.ALIGN_MAP[alignment]
        
        pPr = paragraph._p.get_or_add_pPr()
        
        spacing_before = style.get("spacing_before")
        spacing_after = style.get("spacing_after")
        line_spacing = style.get("line_spacing")
        
        if any([spacing_before, spacing_after, line_spacing]):
            spacing = pPr.find(f"{{{W}}}spacing")
            if spacing is None:
                spacing = OxmlElement("w:spacing")
                pPr.append(spacing)
            
            if spacing_before:
                spacing.set(qn("w:before"), str(spacing_before))
            if spacing_after:
                spacing.set(qn("w:after"), str(spacing_after))
            if line_spacing:
                spacing.set(qn("w:line"), str(line_spacing))
        
        first_line_indent = style.get("first_line_indent")
        left_indent = style.get("left_indent")
        
        if any([first_line_indent, left_indent]):
            ind = pPr.find(f"{{{W}}}ind")
            if ind is None:
                ind = OxmlElement("w:ind")
                pPr.append(ind)
            
            if first_line_indent:
                ind.set(qn("w:firstLine"), str(first_line_indent))
            if left_indent:
                ind.set(qn("w:left"), str(left_indent))
    
    def _restore_paragraph(self, element: Dict[str, Any]):
        content = element.get("content", "")
        style = element.get("style", {})
        elem_type = element.get("type", "body")
        
        paragraph = self.doc.add_paragraph()
        
        if elem_type.startswith("heading"):
            try:
                level = int(elem_type.replace("heading", ""))
                heading_style = self.HEADING_STYLE_MAP.get(level)
                if heading_style:
                    paragraph.style = heading_style
            except (ValueError, KeyError):
                pass
        
        self._apply_paragraph_style(paragraph, style)
        
        paragraph.add_run(content)
    
    def _restore_table(self, element: Dict[str, Any]):
        content = element.get("content", [])
        style = element.get("style", {})
        
        if not content:
            return
        
        rows = len(content)
        cols = max(len(r) for r in content) if content else 0
        
        table = self.doc.add_table(rows=rows, cols=cols)
        table.style = "Table Grid"
        
        for row_idx, row_data in enumerate(content):
            for col_idx, cell_text in enumerate(row_data):
                if col_idx < cols:
                    cell = table.rows[row_idx].cells[col_idx]
                    cell.text = str(cell_text)
    
    def _restore_image(self, element: Dict[str, Any]):
        content = element.get("content", {})
        style = element.get("style", {})
        
        base64_str = content.get("base64", "")
        caption = content.get("caption", "")
        
        if not base64_str:
            paragraph = self.doc.add_paragraph()
            paragraph.add_run("[图片占位符]")
            return
        
        try:
            image_bytes = base64.b64decode(base64_str)
            image_stream = io.BytesIO(image_bytes)
            
            paragraph = self.doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            run = paragraph.add_run()
            
            width_inches = style.get("width_inches")
            if width_inches:
                run.add_picture(image_stream, width=Inches(width_inches))
            else:
                run.add_picture(image_stream)
            
            if caption:
                caption_para = self.doc.add_paragraph()
                caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                caption_para.add_run(caption)
                
        except Exception as e:
            print(f"还原图片时出错: {e}")
            traceback.print_exc()
            paragraph = self.doc.add_paragraph()
            paragraph.add_run(f"[图片还原失败: {e}]")
    
    def _restore_formula(self, element: Dict[str, Any]):
        content = element.get("content", {})
        style = element.get("style", {})
        
        formula_type = style.get("formula_type", "omml")
        
        if formula_type == "ole":
            self._restore_ole_formula(element)
            return
        
        omml_str = content.get("omml", "")
        label = content.get("label", "")
        
        if omml_str:
            try:
                omml_elem = etree.fromstring(omml_str.encode())
                
                p = OxmlElement("w:p")
                pPr = OxmlElement("w:pPr")
                p.append(pPr)
                
                jc = OxmlElement("w:jc")
                jc.set(qn("w:val"), "center")
                pPr.append(jc)
                
                p.append(omml_elem)
                
                if label:
                    run = OxmlElement("w:r")
                    t = OxmlElement("w:t")
                    t.text = f"    {label}"
                    run.append(t)
                    p.append(run)
                
                body = self.doc.element.body
                sectPr = body.find(f"{{{W}}}sectPr")
                if sectPr is not None:
                    sectPr.addprevious(p)
                else:
                    body.append(p)
                    
            except Exception as e:
                print(f"还原公式时出错: {e}")
                traceback.print_exc()
                paragraph = self.doc.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.add_run(f"[公式] {label}")
        else:
            paragraph = self.doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.add_run(f"[公式] {label}")
    
    def _restore_ole_formula(self, element: Dict[str, Any]):
        content = element.get("content", {})
        style = element.get("style", {})
        
        ole_base64 = content.get("ole_base64", "")
        image_base64 = content.get("image_base64", "")
        label = content.get("label", "")
        prog_id = content.get("prog_id", "Equation.3")
        width_pt = style.get("width_pt", 50)
        height_pt = style.get("height_pt", 20)
        
        if not ole_base64:
            paragraph = self.doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.add_run(f"[公式占位符] {label}")
            return
        
        try:
            ole_bytes = base64.b64decode(ole_base64)
            image_bytes = base64.b64decode(image_base64) if image_base64 else None
            
            self._ole_counter += 1
            self._shape_counter += 1
            
            ole_rid, image_rid = self._add_ole_parts(ole_bytes, image_bytes)
            
            shape_id = f"_x0000_i{1025 + self._shape_counter - 1}"
            object_id = f"_{int(1468075725 + self._ole_counter)}"
            
            p = self._create_ole_paragraph(ole_rid, image_rid, shape_id, object_id, 
                                           prog_id, width_pt, height_pt, label)
            
            body = self.doc.element.body
            sectPr = body.find(f"{{{W}}}sectPr")
            if sectPr is not None:
                sectPr.addprevious(p)
            else:
                body.append(p)
                
        except Exception as e:
            print(f"还原OLE公式时出错: {e}")
            traceback.print_exc()
            paragraph = self.doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.add_run(f"[公式还原失败: {e}] {label}")
    
    def _add_ole_parts(self, ole_bytes: bytes, image_bytes: bytes = None) -> tuple:
        ole_rid = self._add_ole_embedded_part(ole_bytes)
        image_rid = self._add_ole_image_part(image_bytes)
        return ole_rid, image_rid
    
    def _get_next_rid(self) -> str:
        rels = self.doc.part.rels
        rids = [r.rId for r in rels.values()]
        return f"rId{max([int(r[3:]) for r in rids if r.startswith('rId')], default=0) + 1}"
    
    def _add_ole_embedded_part(self, ole_bytes: bytes) -> str:
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        
        partname = PackURI(f"/word/embeddings/oleObject{self._ole_counter}.bin")
        content_type = "application/vnd.ms-office.activeX"
        
        part = Part(partname, content_type, ole_bytes)
        
        new_rid = self._get_next_rid()
        rel = self.doc.part.rels.add_relationship(RT.OLE_OBJECT, part, new_rid)
        
        return new_rid
    
    def _add_ole_image_part(self, image_bytes: bytes = None) -> str:
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        
        if image_bytes is None:
            image_bytes = self._create_ole_placeholder_image()
        
        partname = PackURI(f"/word/media/oleImage{self._ole_counter}.wmf")
        content_type = "image/x-wmf"
        
        part = Part(partname, content_type, image_bytes)
        
        new_rid = self._get_next_rid()
        rel = self.doc.part.rels.add_relationship(RT.IMAGE, part, new_rid)
        
        return new_rid
    
    def _create_ole_placeholder_image(self) -> bytes:
        wmf_header = bytes([
            0xD7, 0xCD, 0xC6, 0x9A,
            0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00,
            0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
        ])
        return wmf_header
    
    def _create_ole_paragraph(self, ole_rid: str, image_rid: str, shape_id: str, 
                              object_id: str, prog_id: str, width_pt: float, 
                              height_pt: float, label: str) -> etree._Element:
        p = etree.Element(f"{{{W}}}p")
        
        pPr = etree.SubElement(p, f"{{{W}}}pPr")
        
        jc = etree.SubElement(pPr, f"{{{W}}}jc")
        jc.set(f"{{{W}}}val", "center")
        
        r = etree.SubElement(p, f"{{{W}}}r")
        
        obj = etree.SubElement(r, f"{{{W}}}object")
        
        shape = etree.SubElement(obj, f"{{{V_NS}}}shape")
        shape.set("id", shape_id)
        shape.set(f"{{{O_NS}}}spt", "75")
        shape.set("type", "#_x0000_t75")
        shape.set("style", f"height:{height_pt}pt;width:{width_pt}pt;")
        shape.set(f"{{{O_NS}}}ole", "t")
        shape.set("filled", "f")
        shape.set(f"{{{O_NS}}}preferrelative", "t")
        shape.set("stroked", "f")
        shape.set("coordsize", "21600,21600")
        
        etree.SubElement(shape, f"{{{V_NS}}}path")
        
        fill = etree.SubElement(shape, f"{{{V_NS}}}fill")
        fill.set("on", "f")
        fill.set("focussize", "0,0")
        
        stroke = etree.SubElement(shape, f"{{{V_NS}}}stroke")
        stroke.set("on", "f")
        
        imagedata = etree.SubElement(shape, f"{{{V_NS}}}imagedata")
        imagedata.set(f"{{{R_NS}}}id", image_rid)
        imagedata.set(f"{{{O_NS}}}title", "")
        
        lock = etree.SubElement(shape, f"{{{O_NS}}}lock")
        lock.set(f"{{{V_NS}}}ext", "edit")
        lock.set("grouping", "f")
        lock.set("rotation", "f")
        lock.set("text", "f")
        lock.set("aspectratio", "t")
        
        wrap = etree.SubElement(shape, f"{{{W}}}wrap")
        wrap.set("type", "none")
        
        anchorlock = etree.SubElement(shape, f"{{{W}}}anchorlock")
        
        ole_obj = etree.SubElement(obj, f"{{{O_NS}}}OLEObject")
        ole_obj.set("Type", "Embed")
        ole_obj.set("ProgID", prog_id)
        ole_obj.set("ShapeID", shape_id)
        ole_obj.set("DrawAspect", "Content")
        ole_obj.set("ObjectID", object_id)
        ole_obj.set(f"{{{R_NS}}}id", ole_rid)
        
        locked_field = etree.SubElement(ole_obj, f"{{{O_NS}}}LockedField")
        locked_field.text = "false"
        
        if label:
            r_label = etree.SubElement(p, f"{{{W}}}r")
            t_label = etree.SubElement(r_label, f"{{{W}}}t")
            t_label.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t_label.text = f"    {label}"
        
        return p
    
    def restore(self) -> Document:
        for element in self.elements:
            elem_type = element.get("type", "body")
            
            if elem_type.startswith("heading"):
                self._restore_paragraph(element)
            elif elem_type == "body":
                self._restore_paragraph(element)
            elif elem_type == "table":
                self._restore_table(element)
            elif elem_type == "image":
                self._restore_image(element)
            elif elem_type == "formula":
                self._restore_formula(element)
            else:
                self._restore_paragraph(element)
        
        return self.doc
    
    def save(self, output_path: str):
        self.restore()
        self.doc.save(output_path)


def restore_docx(elements: List[Dict[str, Any]], output_docx_path: str) -> Document:
    """
    从JSON元素列表还原docx文档
    
    Args:
        elements: JSON元素列表
        output_docx_path: 输出docx文件路径
    
    Returns:
        还原后的Document对象
    """
    restorer = DocxRestorer(elements)
    restorer.save(output_docx_path)
    return restorer.doc


def restore_from_json(json_path: str, output_docx_path: str) -> Document:
    """
    从JSON文件还原docx文档
    
    Args:
        json_path: JSON文件路径
        output_docx_path: 输出docx文件路径
    
    Returns:
        还原后的Document对象
    """
    with open(json_path, "r", encoding="utf-8") as f:
        elements = json.load(f)
    
    return restore_docx(elements, output_docx_path)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 3:
        print("用法: python docx_restorer.py <JSON文件路径> <输出docx路径>")
        sys.exit(1)
    
    json_path = sys.argv[1]
    output_path = sys.argv[2]
    
    try:
        restore_from_json(json_path, output_path)
        print(f"还原完成，文档已保存到: {output_path}")
    except Exception as e:
        print(f"还原失败: {e}")
        traceback.print_exc()
