"""
样式回填器
将编辑后的文本样式应用到完整的JSON数据中
"""
import json
import traceback
from typing import List, Dict, Any
from pathlib import Path


def find_toc_element(full_data: List[Dict[str, Any]], heading_content: str) -> Dict[str, Any] | None:
    """
    根据标题内容查找对应的目录项
    
    Args:
        full_data: 完整的JSON数据
        heading_content: 标题内容
    
    Returns:
        匹配的目录项元素，如果没有找到则返回None
    """
    for element in full_data:
        content = element.get("content", "")
        if isinstance(content, str) and "\t" in content:
            toc_text = content.split("\t")[0].strip()
            if toc_text == heading_content.strip():
                return element
    return None


def backfill_styles(edited_data: List[Dict[str, Any]], full_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    将编辑后的样式回填到完整数据中
    
    Args:
        edited_data: 编辑后的文本数据（包含id, content, type）
        full_data: 完整的原始JSON数据
    
    Returns:
        更新后的完整数据
    """
    id_to_element = {e.get("id"): e for e in full_data}
    
    for edited_elem in edited_data:
        elem_id = edited_elem.get("id")
        new_type = edited_elem.get("type", "body")
        
        if elem_id not in id_to_element:
            continue
        
        original_elem = id_to_element[elem_id]
        original_type = original_elem.get("type", "body")
        
        if new_type.startswith("heading"):
            original_elem["type"] = new_type
            
            level = new_type.replace("heading", "")
            heading_content = edited_elem.get("content", "")
            
            toc_elem = find_toc_element(full_data, heading_content)
            if toc_elem:
                toc_elem["type"] = f"toc{level}"
        
        elif new_type == "normal":
            if original_type.startswith("heading"):
                original_elem["type"] = "body"
                
                heading_content = edited_elem.get("content", "")
                toc_elem = find_toc_element(full_data, heading_content)
                if toc_elem and toc_elem.get("type", "").startswith("toc"):
                    toc_elem["type"] = "body"
    
    return full_data


def backfill_from_files(edited_json_path: str, full_json_path: str, output_path: str = None) -> List[Dict[str, Any]]:
    """
    从文件读取数据并执行样式回填
    
    Args:
        edited_json_path: 编辑后的JSON文件路径
        full_json_path: 完整的原始JSON文件路径
        output_path: 输出JSON文件路径（可选）
    
    Returns:
        更新后的完整数据
    """
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


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 3:
        print("用法: python docx_style_backfill.py <编辑后的JSON路径> <完整JSON路径> [输出JSON路径]")
        sys.exit(1)
    
    edited_path = sys.argv[1]
    full_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else None
    
    try:
        updated_data = backfill_from_files(edited_path, full_path, output_path)
        
        heading_count = sum(1 for e in updated_data if e.get("type", "").startswith("heading"))
        toc_count = sum(1 for e in updated_data if e.get("type", "").startswith("toc"))
        
        if output_path:
            print(f"回填完成，已保存到: {output_path}")
            print(f"标题元素: {heading_count} 个")
            print(f"目录元素: {toc_count} 个")
        else:
            print(json.dumps(updated_data, ensure_ascii=False, indent=2))
            
    except Exception as e:
        print(f"回填失败: {e}")
        traceback.print_exc()
