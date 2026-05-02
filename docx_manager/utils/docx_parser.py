"""
DOCX文档样式解析器
将docx文档解析为JSON格式，包含样式信息
"""
import json
import base64
import traceback
import re
from typing import Dict, List, Any, Optional
from pathlib import Path
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
V_NS = "urn:schemas-microsoft-com:vml"

_LABEL_RE       = re.compile(r'\(\d+-\d+\)')
_MODE1_RE       = re.compile(r'^图\s+\d+-\d+\s{2}')   # 图 X-X  内容（独立图题）
_MODE2_RE       = re.compile(r'^\([a-zA-Z]\)\s')        # (n) 内容（子图题）
_FIG_CAPTION_RE = re.compile(r'^图\s+\d+-\d+\s{2}|^\([a-zA-Z]\)\s')


def _split_mode2(text: str) -> list[str]:
    """将 '(a) xxx; (b) yyy' 等 Mode2 段落拆分为独立子图题列表。"""
    markers = list(re.finditer(r'\([a-zA-Z]\)\s', text))
    if not markers:
        return [text]
    result = []
    for i, m in enumerate(markers):
        start = m.start()
        end = markers[i + 1].start() if i + 1 < len(markers) else len(text)
        part = text[start:end].strip().rstrip(';').strip()
        if part:
            result.append(part)
    return result


class DocxParser:
    """DOCX文档样式解析器"""
    
    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.doc = Document(doc_path)
        self.elements: List[Dict[str, Any]] = []
        self.citations: List[Dict[str, Any]] = []
        self._element_id = 0
    
    def _get_next_id(self) -> str:
        self._element_id += 1
        return f"elem_{self._element_id}"
    
    def _extract_superscript_citations(self, paragraph: Paragraph, element_id: str) -> List[Dict[str, Any]]:
        """
        提取段落中格式为上标的文献引用 [x]
        
        Args:
            paragraph: 段落对象
            element_id: 当前元素ID
        
        Returns:
            引用列表 [{"ref_id": 1, "before": "...", "after": "...", "element_id": "elem_1"}, ...]
        """
        citations = []
        
        char_formats = []
        for run in paragraph.runs:
            is_super = (run.font.superscript is True)
            
            for char in run.text:
                char_formats.append((char, is_super))
        
        if not char_formats:
            return citations
        
        full_text = "".join([c[0] for c in char_formats])
        
        for match in re.finditer(r'\[(\d+)\]', full_text):
            match_start, match_end = match.span()
            num_start, num_end = match.span(1)
            
            is_superscript = any(char_formats[i][1] for i in range(num_start, num_end))
            
            if is_superscript:
                before_text = full_text[:match_start]
                after_text = full_text[match_end:]
                
                before_context = before_text[-50:] if len(before_text) > 50 else before_text
                after_context = after_text[:50] if len(after_text) > 50 else after_text
                
                citations.append({
                    "ref_id": int(match.group(1)),
                    "before": before_context.strip(),
                    "after": after_context.strip(),
                    "element_id": element_id,
                    "full_match": match.group(0)
                })
        
        return citations
    
    def _get_paragraph_style(self, paragraph: Paragraph) -> Dict[str, Any]:
        style_info = {}
        
        style_name = paragraph.style.name if paragraph.style else "Normal"
        style_info["style_name"] = style_name
        
        if paragraph.alignment is not None:
            align_map = {
                0: "left",
                1: "center", 
                2: "right",
                3: "justify"
            }
            style_info["alignment"] = align_map.get(paragraph.alignment, "left")
        
        pPr = paragraph._p.find(f"{{{W}}}pPr")
        if pPr is not None:
            spacing = pPr.find(f"{{{W}}}spacing")
            if spacing is not None:
                before = spacing.get(qn("w:before"))
                after = spacing.get(qn("w:after"))
                line = spacing.get(qn("w:line"))
                if before:
                    style_info["spacing_before"] = int(before)
                if after:
                    style_info["spacing_after"] = int(after)
                if line:
                    style_info["line_spacing"] = int(line)
            
            ind = pPr.find(f"{{{W}}}ind")
            if ind is not None:
                first_line = ind.get(qn("w:firstLine"))
                left = ind.get(qn("w:left"))
                if first_line:
                    style_info["first_line_indent"] = int(first_line)
                if left:
                    style_info["left_indent"] = int(left)
        
        style_info["is_heading"] = False
        style_info["heading_level"] = 0
        
        style_lower = style_name.lower()
        if style_lower.startswith("heading"):
            try:
                level = int(''.join(filter(str.isdigit, style_name)))
                style_info["is_heading"] = True
                style_info["heading_level"] = level
            except (ValueError, AttributeError):
                pass
        elif "标题" in style_name:
            import re
            match = re.search(r'标题\s*(\d+)', style_name)
            if match:
                level = int(match.group(1))
                style_info["is_heading"] = True
                style_info["heading_level"] = level
        
        return style_info
    
    def _parse_paragraph(self, paragraph: Paragraph) -> Optional[Dict[str, Any]]:
        """仅处理普通文本段落，公式与图片由 parse() 在上层分流。"""
        text = paragraph.text.strip()
        if not text:
            return None
        style_info = self._get_paragraph_style(paragraph)
        element_type = "body"
        if style_info["is_heading"]:
            element_type = f"heading{style_info['heading_level']}"
        element_id = self._get_next_id()
        citations = self._extract_superscript_citations(paragraph, element_id)
        if citations:
            self.citations.extend(citations)
        return {
            "id": element_id,
            "type": element_type,
            "content": text,
            "style": style_info
        }
    
    def _parse_table(self, table: Table) -> Dict[str, Any]:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                row_data.append(cell_text)
            table_data.append(row_data)
        
        style_info = {
            "style_name": "Table",
            "rows": len(table_data),
            "cols": max(len(r) for r in table_data) if table_data else 0
        }
        
        tbl = table._tbl
        tblPr = tbl.find(f"{{{W}}}tblPr")
        if tblPr is not None:
            tblW = tblPr.find(f"{{{W}}}tblW")
            if tblW is not None:
                w = tblW.get(qn("w:w"))
                if w:
                    style_info["width"] = int(w)
        
        return {
            "id": self._get_next_id(),
            "type": "table",
            "content": table_data,
            "style": style_info
        }
    
    def _parse_images(self, paragraph: Paragraph) -> list[dict]:
        """提取段落内所有用户插入的内嵌图片，返回 image dict 列表。
        只处理 wp:inline（嵌入型），跳过 wp:anchor（浮动/模板图）。
        caption 留空，由 parse() 按全文序号填入。"""
        results = []
        for run in paragraph.runs:
            for drawing in run._element.findall(".//" + qn("w:drawing")):
                inline = drawing.find(qn("wp:inline"))
                if inline is None:
                    continue  # 跳过浮动型锚点图
                for blip in drawing.findall(".//" + qn("a:blip")):
                    embed_attr = blip.get(qn("r:embed"))
                    if not embed_attr:
                        continue
                    try:
                        image_part = self.doc.part.related_parts.get(embed_attr)
                        if not image_part:
                            continue
                        image_bytes = image_part.blob
                        ext = image_part.content_type.split('/')[-1]
                        if ext == 'jpeg':
                            ext = 'jpg'
                        base64_str = base64.b64encode(image_bytes).decode('utf-8')
                        width_inches = None
                        extent = inline.find(qn("wp:extent"))
                        if extent is not None:
                            cx = extent.get("cx")
                            if cx:
                                width_inches = int(cx) / 914400
                        results.append({
                            "id": self._get_next_id(),
                            "type": "image",
                            "base64": base64_str,
                            "caption": "",
                            "width": width_inches,
                            "height": width_inches,
                            "position": "center",
                        })
                    except Exception as e:
                        print(f"解析图片时出错: {e}")
                        traceback.print_exc()
            for pict in run._element.findall(".//" + qn("w:pict")):
                for imagedata in pict.findall(".//{" + V_NS + "}imagedata"):
                    embed_attr = imagedata.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                    )
                    if not embed_attr:
                        continue
                    try:
                        image_part = self.doc.part.related_parts.get(embed_attr)
                        if not image_part:
                            continue
                        image_bytes = image_part.blob
                        ext = image_part.content_type.split('/')[-1]
                        if ext == 'jpeg':
                            ext = 'jpg'
                        if ext in ('x-wmf', 'wmf'):
                            continue  # WMF 几乎都是 OLE 公式预览，跳过
                        base64_str = base64.b64encode(image_bytes).decode('utf-8')
                        results.append({
                            "id": self._get_next_id(),
                            "type": "image",
                            "base64": base64_str,
                            "caption": "",
                            "width": width_inches,
                            "height": width_inches,
                            "position": "center",
                        })
                    except Exception as e:
                        print(f"解析VML图片时出错: {e}")
                        traceback.print_exc()
        return results
    
    def _extract_surrounding_text(self, paragraph: Paragraph, math_elem) -> tuple:
        before_parts = []
        after_parts = []
        found = False

        for child in paragraph._element:
            if child is math_elem or math_elem in child.iter():
                found = True
                continue
            text = "".join(t.text or "" for t in child.findall(f".//{{{W}}}t"))
            if not text:
                continue
            if not found:
                before_parts.append(text)
            else:
                after_parts.append(text)

        text_before = "".join(before_parts).strip() or None
        text_after = _LABEL_RE.sub("", "".join(after_parts)).strip() or None
        return text_before, text_after

    def _parse_formula(self, paragraph: Paragraph) -> Optional[Dict[str, Any]]:
        omath_paras = paragraph._element.findall(".//" + f"{{{MATH_NS}}}oMathPara")
        
        for omath_para in omath_paras:
            try:
                omml_str = ET.tostring(omath_para, encoding="unicode", method="xml")

                label_match = _LABEL_RE.search(paragraph.text)
                label = label_match.group() if label_match else ""
                text_before, text_after = self._extract_surrounding_text(paragraph, omath_para)

                result = {"type": "formula", "label": label, "omml": omml_str}
                if text_before:
                    result["text_before"] = text_before
                if text_after:
                    result["text_after"] = text_after
                return result
            except Exception as e:
                print(f"解析公式时出错: {e}")
                traceback.print_exc()

        omaths = paragraph._element.findall(".//" + f"{{{MATH_NS}}}oMath")
        if omaths:
            try:
                omml_str = ET.tostring(omaths[0], encoding="unicode", method="xml")

                label_match = _LABEL_RE.search(paragraph.text)
                label = label_match.group() if label_match else ""
                text_before, text_after = self._extract_surrounding_text(paragraph, omaths[0])

                result = {"type": "formula", "label": label, "omml": omml_str, "is_inline": True}
                if text_before:
                    result["text_before"] = text_before
                if text_after:
                    result["text_after"] = text_after
                return result
            except Exception as e:
                print(f"解析公式时出错: {e}")
                traceback.print_exc()
        
        ole_result = self._parse_ole_formula(paragraph)
        if ole_result:
            return ole_result
        
        return None
    
    def _parse_ole_formula(self, paragraph: Paragraph) -> Optional[Dict[str, Any]]:
        O_NS = "urn:schemas-microsoft-com:office:office"
        V_NS = "urn:schemas-microsoft-com:vml"
        
        objects = paragraph._element.findall(".//" + f"{{{W}}}object")
        
        if not objects:
            return None
        
        for obj in objects:
            ole_objects = obj.findall(".//" + f"{{{O_NS}}}OLEObject")
            for ole_obj in ole_objects:
                prog_id = ole_obj.get("ProgID", "")
                if "Equation" in prog_id or "equation" in prog_id.lower():
                    try:
                        embed_id = ole_obj.get(f"{{{qn('r:embed').split('}')[0]}}}id")
                        if not embed_id:
                            embed_id = ole_obj.get(qn("r:id"))
                        if embed_id:
                            ole_part = self.doc.part.related_parts.get(embed_id)
                            if ole_part:
                                ole_bytes = ole_part.blob
                                base64_str = base64.b64encode(ole_bytes).decode('utf-8')
                                
                                label_match = _LABEL_RE.search(paragraph.text)
                                label = label_match.group() if label_match else ""
                                text_before, text_after = self._extract_surrounding_text(paragraph, obj)

                                shape_elem = obj.find(".//" + f"{{{V_NS}}}shape")
                                width_pt = None
                                height_pt = None
                                image_base64 = None
                                
                                if shape_elem is not None:
                                    style = shape_elem.get("style", "")
                                    if "width:" in style:
                                        width_match = re.search(r'width:([\d.]+)pt', style)
                                        if width_match:
                                            width_pt = float(width_match.group(1))
                                    if "height:" in style:
                                        height_match = re.search(r'height:([\d.]+)pt', style)
                                        if height_match:
                                            height_pt = float(height_match.group(1))
                                    
                                    imagedata = shape_elem.find(f".//{{{V_NS}}}imagedata")
                                    if imagedata is not None:
                                        image_rid = imagedata.get(f"{{{qn('r:id').split('}')[0]}}}id")
                                        if not image_rid:
                                            image_rid = imagedata.get(qn("r:id"))
                                        if image_rid:
                                            image_part = self.doc.part.related_parts.get(image_rid)
                                            if image_part:
                                                image_bytes = image_part.blob
                                                image_base64 = base64.b64encode(image_bytes).decode('utf-8')
                                
                                result = {
                                    "type": "formula",
                                    "label": label,
                                    "omml": "",
                                    "ole_base64": base64_str,
                                    "prog_id": prog_id
                                }

                                if text_before:
                                    result["text_before"] = text_before
                                if text_after:
                                    result["text_after"] = text_after
                                if image_base64:
                                    result["image_base64"] = image_base64
                                if width_pt is not None:
                                    result["width_pt"] = width_pt
                                if height_pt is not None:
                                    result["height_pt"] = height_pt

                                return result
                    except Exception as e:
                        print(f"解析OLE公式时出错: {e}")
                        traceback.print_exc()
        
        return None
    
    def _prescan_captions(self) -> tuple[list[str], set[int]]:
        """全文预扫描，收集所有可匹配图题的段落。

        规则：
          Mode 2  (n) xxx          → 子图题，按 (字母) 拆分为独立条目后收录
          Mode 1  图 X-X  xxx      → 独立图题，收录
          Mode 1 紧跟 Mode 2 之后  → 总图题，不收录（保留为正文）
        空段落不重置状态机。
        返回：(有序图题列表, 需跳过的 body 子元素索引集合)
        """
        captions: list[str] = []
        skip_indices: set[int] = set()   # body 子元素的位置索引，与 id() 无关
        prev_mode2 = False

        body_children = list(self.doc.element.body)
        for idx, element in enumerate(body_children):
            if element.tag != f"{{{W}}}p":
                prev_mode2 = False
                continue
            text = Paragraph(element, self.doc).text.strip()
            if not text:
                continue  # 空段落不改变状态

            if _MODE2_RE.match(text):
                captions.extend(_split_mode2(text))   # 拆分后逐条收录
                skip_indices.add(idx)
                prev_mode2 = True
            elif _MODE1_RE.match(text):
                if not prev_mode2:
                    captions.append(text)
                    skip_indices.add(idx)
                # prev_mode2=True 时：总图题，不收录，不跳过（保留为正文）
                prev_mode2 = False
            else:
                prev_mode2 = False

        return captions, skip_indices

    def parse(self) -> List[Dict[str, Any]]:
        self.elements = []
        self.citations = []

        # Phase 1: 全文预扫描图题
        captions, caption_skip_indices = self._prescan_captions()
        caption_cursor = 0

        # Phase 2: 主循环，用索引判断跳过（避免 lxml proxy id() 不稳定的问题）
        body_children = list(self.doc.element.body)
        for idx, element in enumerate(body_children):
            if idx in caption_skip_indices:
                continue

            if element.tag == f"{{{W}}}p":
                paragraph = Paragraph(element, self.doc)

                formula_result = self._parse_formula(paragraph)
                if formula_result:
                    self.elements.append(formula_result)
                    continue

                images = self._parse_images(paragraph)
                if images:
                    for img in images:
                        img['caption'] = (
                            captions[caption_cursor]
                            if caption_cursor < len(captions)
                            else ""
                        )
                        caption_cursor += 1
                        self.elements.append(img)
                    continue

                result = self._parse_paragraph(paragraph)
                if result:
                    self.elements.append(result)

            elif element.tag == f"{{{W}}}tbl":
                table = Table(element, self.doc.element.body)
                self.elements.append(self._parse_table(table))

        return self.elements
    
    def get_citations(self) -> List[Dict[str, Any]]:
        return self.citations
    
    def save_citations(self, output_path: str) -> None:
        if self.citations:
            Path(output_path).write_text(
                json.dumps(self.citations, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
    
    def to_json(self, output_path: Optional[str] = None, citations_path: Optional[str] = None) -> str:
        if not self.elements:
            self.parse()
        
        json_str = json.dumps(self.elements, ensure_ascii=False, indent=2)
        
        if output_path:
            Path(output_path).write_text(json_str, encoding="utf-8")
        
        if citations_path and self.citations:
            self.save_citations(citations_path)
        
        return json_str


def parse_docx(doc_path: str, output_json_path: Optional[str] = None, citations_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    解析docx文档为JSON格式
    
    Args:
        doc_path: docx文档路径
        output_json_path: 输出JSON文件路径（可选）
        citations_path: 引用配置文件路径（可选）
    
    Returns:
        解析后的元素列表
    """
    parser = DocxParser(doc_path)
    elements = parser.parse()
    
    if output_json_path:
        parser.to_json(output_json_path, citations_path)
    elif citations_path and parser.citations:
        parser.save_citations(citations_path)
    
    return elements


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("用法: python docx_parser.py <docx文件路径> [输出JSON路径]")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        parser = DocxParser(doc_path)
        elements = parser.parse()
        
        if output_path:
            parser.to_json(output_path)
            print(f"解析完成，结果已保存到: {output_path}")
        else:
            print(json.dumps(elements, ensure_ascii=False, indent=2))
    except Exception as e:
        print(f"解析失败: {e}")
        traceback.print_exc()
