"""
文本提取器
从docx_parser生成的JSON中提取文本类型数据
"""
import json
import traceback
from typing import List, Dict, Any
from pathlib import Path


TEXT_TYPES = {"body", "heading1", "heading2", "heading3"}


def extract_text_elements(elements: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """
    从解析后的元素列表中提取文本类型元素
    
    Args:
        elements: docx_parser解析后的元素列表
    
    Returns:
        只包含id和content的文本元素列表
    """
    text_elements = []
    
    for element in elements:
        elem_type = element.get("type", "")
        
        if elem_type in TEXT_TYPES:
            content = element.get("content", "")
            
            if content and isinstance(content, str):
                text_elements.append({
                    "id": element.get("id", ""),
                    "content": content.strip()
                })
    
    return text_elements


def extract_text_from_json(json_path: str, output_path: str = None) -> List[Dict[str, str]]:
    """
    从JSON文件中提取文本元素
    
    Args:
        json_path: 输入JSON文件路径
        output_path: 输出JSON文件路径（可选）
    
    Returns:
        文本元素列表
    """
    with open(json_path, 'r', encoding='utf-8') as f:
        elements = json.load(f)
    
    text_elements = extract_text_elements(elements)
    
    if output_path:
        result = {"text_elements": text_elements}
        Path(output_path).write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding='utf-8'
        )
    
    return text_elements


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("用法: python text_extractor.py <json文件路径> [输出JSON路径]")
        sys.exit(1)
    
    json_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        text_elements = extract_text_from_json(json_path, output_path)
        
        if output_path:
            print(f"提取完成，共 {len(text_elements)} 个文本元素，已保存到: {output_path}")
        else:
            print(json.dumps(text_elements, ensure_ascii=False, indent=2))
            print(f"\n共 {len(text_elements)} 个文本元素")
            
    except Exception as e:
        print(f"提取失败: {e}")
        traceback.print_exc()
