from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def _iter_elements_by_tag(element, tag_local_name):
    for child in element.iter():
        if child.tag.endswith("}" + tag_local_name):
            yield child

def _find_child(element, tag):
    return element.find(f"{{{W}}}{tag}")

def debug_docx_structure(doc_path):
    doc = Document(doc_path)
    print(f"=== Debugging {doc_path} ===")
    
    element_count = 0
    paragraph_count = 0
    table_count = 0
    image_count = 0
    
    for element in doc.element.body:
        element_count += 1
        local = element.tag.split("}")[-1]
        
        if local == "p":
            paragraph_count += 1
            paragraph = Paragraph(element, doc)
            text = paragraph.text.strip()
            
            if text:
                print(f"段落 {paragraph_count}: {text[:50]}...")
            
            # 检查是否包含图片
            has_image = False
            for run in paragraph.runs:
                for drawing in _iter_elements_by_tag(run._element, "drawing"):
                    print(f"  -> 发现 drawing 元素")
                    # 检查 drawing 下的子元素
                    for child in drawing.iter():
                        tag_name = child.tag.split("}")[-1]
                        print(f"     子元素: {tag_name}")
                        if tag_name in ["blipFill", "pic", "blip"]:
                            has_image = True
            
            if has_image:
                image_count += 1
                print(f"  -> 发现图片！")
        
        elif local == "tbl":
            table_count += 1
            print(f"表格 {table_count}")
    
    print(f"\n总计:")
    print(f"  元素数: {element_count}")
    print(f"  段落数: {paragraph_count}")
    print(f"  表格数: {table_count}")
    print(f"  图片数: {image_count}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        doc_path = sys.argv[1]
    else:
        doc_path = "data/output.docx"
    
    debug_docx_structure(doc_path)