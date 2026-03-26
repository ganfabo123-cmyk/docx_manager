import json
import traceback
from typing import List, Dict, Any

try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("python-docx not installed, docx generation will be disabled")

def parse_document_content(content: str) -> List[Dict[str, Any]]:
    """
    Parse document content into text blocks
    """
    try:
        blocks = []
        for i, line in enumerate(content.split('\n'), 1):
            if line.strip():
                blocks.append({
                    'id': i,
                    'content': line.strip()
                })
        return blocks
    except Exception as e:
        print(f"Error parsing document: {e}")
        print(traceback.format_exc())
        return []

def identify_short_blocks(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Identify short text blocks (content length < 30)
    """
    try:
        return [block for block in blocks if len(block['content']) < 30]
    except Exception as e:
        print(f"Error identifying short blocks: {e}")
        print(traceback.format_exc())
        return []

def generate_llm_prompt(blocks: List[Dict[str, Any]]) -> str:
    """
    Generate prompt for LLM to analyze styles
    """
    try:
        prompt = f"""Analyze the following short text blocks and determine their style type:

{json.dumps(blocks, ensure_ascii=False, indent=2)}

Please identify which blocks are:
- Level 1 heading (h1)
- Level 2 heading (h2)
- Level 3 heading (h3)
- Table of contents item (toc1, toc2, toc3)
- Other (leave as paragraph)

Please return a JSON array where each object has:
- id: the same as input
- type: the determined style type
- content: the original content
"""
        return prompt
    except Exception as e:
        print(f"Error generating LLM prompt: {e}")
        print(traceback.format_exc())
        return ""

def save_parsed_styles(analysis: List[Dict[str, Any]], output_path: str) -> None:
    """
    Save parsed styles to JSON file
    """
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(analysis, f, ensure_ascii=False, indent=2)
        print(f"Parsed styles saved to {output_path}")
    except Exception as e:
        print(f"Error saving parsed styles: {e}")
        print(traceback.format_exc())

def generate_styled_content(parsed_styles: List[Dict[str, Any]]) -> str:
    """
    Generate styled content from parsed styles
    """
    try:
        styled_content = ""
        for item in parsed_styles:
            if item['type'] == 'h1':
                styled_content += f"# {item['content']}\n\n"
            elif item['type'] == 'h2':
                styled_content += f"## {item['content']}\n\n"
            elif item['type'] == 'h3':
                styled_content += f"### {item['content']}\n\n"
            elif item['type'] in ['toc1', 'toc2', 'toc3']:
                styled_content += f"{item['content']}\n"
            else:
                styled_content += f"{item['content']}\n\n"
        return styled_content
    except Exception as e:
        print(f"Error generating styled content: {e}")
        print(traceback.format_exc())
        return ""

def generate_docx_document(parsed_styles: List[Dict[str, Any]], output_path: str) -> bool:
    """
    Generate docx document from parsed styles
    """
    try:
        doc = Document()
        
        for item in parsed_styles:
            if item['type'] == 'h1':
                # Level 1 heading
                para = doc.add_heading(item['content'], level=1)
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif item['type'] == 'h2':
                # Level 2 heading
                para = doc.add_heading(item['content'], level=2)
            elif item['type'] == 'h3':
                # Level 3 heading
                para = doc.add_heading(item['content'], level=3)
            elif item['type'] in ['toc1', 'toc2', 'toc3']:
                # Table of contents item
                para = doc.add_paragraph(item['content'])
            else:
                # Regular paragraph
                para = doc.add_paragraph(item['content'])
        
        doc.save(output_path)
        print(f"Generated docx document saved to: {output_path}")
        return True
    except Exception as e:
        print(f"Error generating docx document: {e}")
        print(traceback.format_exc())
        return False