"""
文件解析器
提供各种文件格式的解析、转换接口
"""
import json
import re
import traceback
from typing import Dict, List, Any, Optional
from pathlib import Path
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from utils.docx_parser import DocxParser
from utils.text_extractor import extract_text_elements
from utils.docx_style_backfill import backfill_styles
from utils.docx_restorer import DocxRestorer


def remove_markdown_symbols(text: str) -> str:
    """
    去除文本中的 Markdown 符号
    
    Args:
        text: 包含 Markdown 符号的文本
    
    Returns:
        去除 Markdown 符号后的纯文本
    """
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)
    text = re.sub(r'_(.+?)_', r'\1', text)
    text = re.sub(r'~~(.+?)~~', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    text = re.sub(r'^>\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'^[-*+]\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'^\d+\.\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'^---+\s*$', '', text, flags=re.MULTILINE)
    text = re.sub(r'^\*\*\*+\s*$', '', text, flags=re.MULTILINE)
    text = re.sub(r'^```[\s\S]*?^```', '', text, flags=re.MULTILINE)
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    return text.strip()


def parse_text_to_elements(text: str, remove_md: bool = True) -> List[Dict[str, Any]]:
    """
    将纯文本解析为文档元素列表
    
    Args:
        text: 输入文本（可能包含 Markdown 格式）
        remove_md: 是否去除 Markdown 符号
    
    Returns:
        文档元素列表 [{"id": "elem_1", "type": "body", "content": "xxx", "style": {...}}, ...]
    """
    try:
        if remove_md:
            text = remove_markdown_symbols(text)
        
        elements = []
        lines = text.split('\n')
        
        for i, line in enumerate(lines, 1):
            line = line.strip()
            if not line:
                continue
            
            element = {
                "id": f"elem_{i}",
                "type": "body",
                "content": line,
                "style": {
                    "style_name": "Normal",
                    "alignment": "left"
                }
            }
            elements.append(element)
        
        return elements
    except Exception as e:
        print(f"解析文本失败: {e}")
        traceback.print_exc()
        return []


def parse_text_to_json(text: str, output_json_path: str = None, remove_md: bool = True) -> Optional[List[Dict[str, Any]]]:
    """
    将纯文本解析为JSON格式并可选保存到文件
    
    Args:
        text: 输入文本
        output_json_path: 输出JSON文件路径（可选）
        remove_md: 是否去除 Markdown 符号
    
    Returns:
        文档元素列表
    """
    try:
        elements = parse_text_to_elements(text, remove_md)
        
        if output_json_path and elements:
            Path(output_json_path).write_text(
                json.dumps(elements, ensure_ascii=False, indent=2),
                encoding='utf-8'
            )
        short_elements = [e for e in elements if len(e.get('content','')) ]
        return short_elements
    except Exception as e:
        print(f"解析文本到JSON失败: {e}")
        traceback.print_exc()
        return None

# 在 file_parser.py 中修改读取逻辑
def parse_txt_file(file_path,remove_md):
    # 尝试多种编码读取
    try:    
        encodings = ['utf-8', 'gbk', 'utf-16', 'ansi']
        content = None
        blocks = []
        
        for enc in encodings:
            try:
                with open(file_path, 'r', encoding=enc) as f:
                    content = f.read()
                print(f"Successfully decoded with {enc}")
                break
            except UnicodeDecodeError:
                continue
                
        if remove_md:
            content = remove_markdown_symbols(content)
            
        for i, line in enumerate(content.split('\n'), 1):
            if line.strip():
                blocks.append({
                    'id': f'elem_{i}',
                    'content': line.strip()
                    })
        return {"text_elements": blocks}
    except Exception as e:
        print(f"Error parsing txt file: {e}")
        traceback.print_exc()
        return None      

            

def parse_txt_file(file_path: str, remove_md: bool = True) -> Optional[Dict[str, List[Dict[str, str]]]]:
    """
    Parse text from txt file into text blocks
    
    Args:
        file_path: txt文件路径
        remove_md: 是否去除 Markdown 符号
    
    Returns:
        {"text_elements": [{"id": "elem_1", "content": "xxx"}, ...]}
    """
    try:
        blocks = []
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        if remove_md:
            content = remove_markdown_symbols(content)
        
        for i, line in enumerate(content.split('\n'), 1):
            if line.strip():
                blocks.append({
                    'id': f'elem_{i}',
                    'content': line.strip()
                })
        return {"text_elements": blocks}
    except Exception as e:
        print(f"Error parsing txt file: {e}")
        traceback.print_exc()
        return None


def parse_docx_to_json(docx_path: str, output_json_path: str = None) -> Optional[List[Dict[str, Any]]]:
    """
    解析DOCX文档为JSON格式
    
    Args:
        docx_path: docx文档路径
        output_json_path: 输出JSON文件路径（可选）
    
    Returns:
        解析后的元素列表
    """
    try:
        parser = DocxParser(docx_path)
        elements = parser.parse()
        
        if output_json_path:
            parser.to_json(output_json_path)
        
        return elements
    except Exception as e:
        print(f"解析DOCX失败: {e}")
        traceback.print_exc()
        return None


def extract_text_from_parsed_json(json_path: str, output_path: str = None, remove_md: bool = True) -> Optional[List[Dict[str, str]]]:
    """
    从解析后的JSON中提取文本元素
    
    Args:
        json_path: 输入JSON文件路径
        output_path: 输出JSON文件路径（可选）
        remove_md: 是否去除 Markdown 符号
    
    Returns:
        文本元素列表 [{"id": "elem_1", "content": "xxx"}, ...]
    """
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            elements = json.load(f)
        
        text_elements = extract_text_elements(elements)
        
        if remove_md:
            for elem in text_elements:
                if 'content' in elem:
                    elem['content'] = remove_markdown_symbols(elem['content'])
        
        if output_path:
            result = {"text_elements": text_elements}
            Path(output_path).write_text(
                json.dumps(result, ensure_ascii=False, indent=2),
                encoding='utf-8'
            )
        
        return text_elements
    except Exception as e:
        print(f"提取文本失败: {e}")
        traceback.print_exc()
        return None


def backfill_styles_to_json(edited_json_path: str, full_json_path: str, output_path: str = None) -> Optional[List[Dict[str, Any]]]:
    """
    将编辑后的样式回填到完整JSON中
    
    Args:
        edited_json_path: 编辑后的JSON文件路径
        full_json_path: 完整的原始JSON文件路径
        output_path: 输出JSON文件路径（可选）
    
    Returns:
        更新后的完整数据
    """
    try:
        with open(edited_json_path, 'r', encoding='utf-8') as f:
            edited_data = json.load(f)
        
        if isinstance(edited_data, dict) and "text_elements" in edited_data:
            edited_data = edited_data["text_elements"]
        
        with open(full_json_path, 'r', encoding='utf-8') as f:
            full_data = json.load(f)
        
        updated_data = backfill_styles(edited_data, full_data)
        
        if output_path:
            Path(output_path).write_text(
                json.dumps(updated_data, ensure_ascii=False, indent=2),
                encoding='utf-8'
            )
        
        return updated_data
    except Exception as e:
        print(f"样式回填失败: {e}")
        traceback.print_exc()
        return None


def restore_docx_from_json(json_path: str, output_docx_path: str) -> bool:
    """
    从JSON还原DOCX文档
    
    Args:
        json_path: JSON文件路径
        output_docx_path: 输出docx文件路径
    
    Returns:
        是否成功
    """
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            elements = json.load(f)
        
        restorer = DocxRestorer(elements)
        restorer.save(output_docx_path)
        
        return True
    except Exception as e:
        print(f"还原DOCX失败: {e}")
        traceback.print_exc()
        return False


def parse_file(file_path: str, output_json_path: str = None, remove_md: bool = True) -> Optional[Dict[str, List[Dict[str, str]]]]:
    """
    统一的文件解析接口，根据文件扩展名自动选择解析方式
    
    Args:
        file_path: 输入文件路径（支持 .txt 和 .docx）
        output_json_path: 输出JSON文件路径（可选，仅对docx有效）
        remove_md: 是否去除 Markdown 符号（仅对 txt 有效）
    
    Returns:
        {"text_elements": [{"id": "elem_1", "content": "xxx"}, ...]}
    """
    try:
        ext = Path(file_path).suffix.lower()
        
        if ext == '.txt':
            return parse_txt_file(file_path, remove_md=remove_md)
        elif ext == '.docx':
            elements = parse_docx_to_json(file_path, output_json_path)
            if elements:
                text_elements = extract_text_elements(elements)
                if remove_md:
                    for elem in text_elements:
                        if 'content' in elem:
                            elem['content'] = remove_markdown_symbols(elem['content'])
                return {"text_elements": text_elements}
            return None
        else:
            print(f"不支持的文件格式: {ext}")
            return None
    except Exception as e:
        print(f"解析文件失败: {e}")
        traceback.print_exc()
        return None
