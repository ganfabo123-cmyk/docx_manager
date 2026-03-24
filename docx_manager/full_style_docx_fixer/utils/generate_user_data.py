import json
import re
from typing import List, Dict, Any, Optional
from pathlib import Path


def roman_to_int(roman: str) -> int:
    roman_numerals = {
        'i': 1, 'v': 5, 'x': 10, 'l': 50, 'c': 100, 'd': 500, 'm': 1000,
        'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000
    }
    
    if not roman:
        return None
    
    for char in roman:
        if char not in roman_numerals:
            return None
    
    result = 0
    prev_value = 0
    for char in reversed(roman):
        value = roman_numerals[char]
        if value < prev_value:
            result -= value
        else:
            result += value
        prev_value = value
    
    return result


def parse_page_number(page_str: str) -> str:
    page_str = page_str.strip()
    
    try:
        int(page_str)
        return page_str
    except ValueError:
        pass
    
    roman_result = roman_to_int(page_str)
    if roman_result is not None:
        return page_str.lower()
    
    return page_str


def extract_toc_entry(toc_item: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    value = toc_item.get('value', '')
    
    match = re.match(r'^(.+?)\t(.+)$', value)
    if not match:
        return None
    
    title = match.group(1).strip()
    page_str = match.group(2).strip()
    page = parse_page_number(page_str)
    
    toc_type = toc_item.get('type', 'toc1')
    level_match = re.search(r'toc(\d)', toc_type)
    level = int(level_match.group(1)) if level_match else 1
    
    return {
        'title': title,
        'level': level,
        'page': page
    }


def is_section_type(item_type: str) -> Optional[str]:
    section_mapping = {
        'abstract': 'abstract',
        'abstract_en': 'abstract_en',
        'conclusion': 'conclusion',
        'acknowledgement': 'acknowledgement',
        'references': 'references'
    }
    return section_mapping.get(item_type)


def convert_section(item: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    section_type = is_section_type(item['type'])
    toc_exclude = config.get('section_toc_exclude', {}).get(section_type, True)
    
    return {
        'type': 'section',
        'section_type': section_type,
        'toc_exclude': toc_exclude,
        'value': item.get('value', '')
    }


def convert_heading(item: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    toc_exclude = config.get('heading_toc_exclude_default', False)
    
    result = {
        'type': item['type'],
        'value': item.get('value', '')
    }
    
    if toc_exclude:
        result['toc_exclude'] = toc_exclude
    
    return result


def convert_body(item: Dict[str, Any]) -> Dict[str, Any]:
    return {
        'type': 'body',
        'value': item.get('value', '')
    }


def convert_table(item: Dict[str, Any]) -> Dict[str, Any]:
    result = {
        'type': 'table',
        'data': item.get('data', [])
    }
    
    if item.get('caption'):
        result['caption'] = item['caption']
    
    return result


def convert_image(item: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    image_defaults = config.get('image_defaults', {})
    
    value = item.get('value', {})
    base64_str = value.get('base', '') if isinstance(value, dict) else ''
    caption = value.get('caption', '') if isinstance(value, dict) else ''
    
    result = {
        'type': 'image',
        'base64': base64_str,
        'ext': image_defaults.get('ext', 'png'),
        'width': image_defaults.get('width', 3.5),
        'align': image_defaults.get('align', 'center')
    }
    
    if caption:
        result['caption'] = caption
    
    return result


def convert_formula(item: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    result = {
        'type': 'formula',
        'omml': item.get('omml', '')
    }
    
    if item.get('label'):
        result['label'] = item['label']
    
    return result


def convert_reference(item: Dict[str, Any]) -> Dict[str, Any]:
    return {
        'type': 'reference',
        'id': item.get('id'),
        'text': item.get('text', '')
    }


def load_config(config_path: str) -> Dict[str, Any]:
    path = Path(config_path)
    if path.exists():
        print("配置文件存在!")
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        print("配置文件不存在")
    return {}


def is_special_section_title(title: str) -> Optional[str]:
    if not title:
        return None
    
    title = title.strip()
    title_no_spaces = title.replace(' ', '')
    title_lower = title.lower()
    
    # 摘要
    if '摘要' in title_no_spaces:
        return 'abstract'
    if '目录' in title_no_spaces:
        return 'toc'
    # Abstract
    elif title_lower == 'abstract':
        return 'abstract_en'
    # 结论
    elif '结论' in title_no_spaces:
        return 'conclusion'
    # 致谢
    elif '致谢' in title_no_spaces:
        return 'acknowledgement'
    # 已发表的学术论文目录
    elif '已发表的学术论文目录' in title_no_spaces:
        return 'publications'
    # 附录
    elif '附录' in title_no_spaces:
        return 'custom'
    # 参考文献
    elif '参考文献' in title_no_spaces:
        return 'references'
    
    return None

def generate_user_data(docx_info: List[Dict[str, Any]], config: Dict[str, Any], extracted_citations: List[Dict[str, Any]] = None) -> Dict[str, Any]:
    content = []
    toc_entries = []
    references = []
    
    # 记录哪些索引已经被合并处理过了，避免重复处理
    processed_indices = set()
    
    for i in range(len(docx_info)):
        # --- 这里可以放你的调试代码，现在它能捕捉到每一个 i 了 ---
        if i == 44:
            # 现在 i=44 一定会进入这里
            print(f"Debug: Processing index {i}, type: {docx_info[i].get('type')}")
            
        if i in processed_indices:
            continue
            
        item = docx_info[i]
        item_type = item.get('type', '')
        
        # 1. 处理 TOC
        if item_type.startswith('toc'):
            toc_entry = extract_toc_entry(item)
            if toc_entry:
                toc_entries.append(toc_entry)
            continue
        
        # 2. 处理 Section 类型
        if is_section_type(item_type):
            content.append(convert_section(item, config))
            continue
            
        # 3. 处理 Heading 及其合并逻辑
        if item_type.startswith('heading'):
            heading_title = item.get('value', '')
            section_type = is_special_section_title(heading_title)
            
            if section_type:
                # 发现特殊章节，开始向后寻找属于该章节的 body
                section_content = []
                j = i + 1
                while j < len(docx_info):
                    next_item = docx_info[j]
                    next_type = next_item.get('type', '')
                    
                    # 遇到下一个标题就停止合并
                    if next_type.startswith('heading'):
                        break
                    
                    if next_type == 'body':
                        section_content.append(next_item.get('value', ''))
                        processed_indices.add(j) # 标记此 body 已被合并
                    j += 1
                
                combined_content = '\n\n'.join(section_content)
                
                # 分发逻辑
                if section_type == 'references':
                    ref_pattern = re.compile(r'\［(\d+)\］(.+)')
                    for body_text in section_content:
                        match = ref_pattern.match(body_text.strip())
                        if match:
                            ref_id = int(match.group(1))
                            ref_text = match.group(2).strip()
                            references.append({"id": ref_id, "text": ref_text})
                elif section_type == 'toc':
                    content.append({
                        'type': 'toc',
                        'toc_title_exclude': True,
                        'title': "目  录"
                    })
                else:
                    section_item = {
                        'type': 'section',
                        'section_type': section_type,
                        'toc_exclude': section_type in ['abstract', 'abstract_en', 'custom'],
                        'value': combined_content
                    }
                    if section_type == 'custom':
                        section_item['title'] = heading_title
                    content.append(section_item)
                continue
            else:
                # 普通标题
                content.append(convert_heading(item, config))
                continue
                
        # 4. 处理其他原子类型
        if item_type == 'body':
            content.append(convert_body(item))
        elif item_type == 'table':
            content.append(convert_table(item))
        elif item_type == 'image':
            content.append(convert_image(item, config))
        elif item_type == 'formula':
            content.append(convert_formula(item, config))
        elif item_type == 'reference':
            content.append(convert_reference(item))
            
    # 组装结果 (保持原样)
    result = {
        '_doc': '由 parse_full_docx 生成的数据转换而来',
        '_tips': {
            '图片_path模式': '"path": "/absolute/path/to/image.png"',
            '图片_base64模式': '"base64": "<base64字符串>", "ext": "png"',
            'toc_exclude': 'true → 标题不进 TOC 域',
            '公式_omml': '直接嵌入 Office Open Math XML',
            '公式_latex': '需要 pip install latex2mathml'
        },
        'page_footer_config': config.get('page_footer_config', []),
        'toc_mode': config.get('toc_mode', 'manual'),
        'toc_entries': toc_entries,
        'content': content
    }
    
    if references:
        result['references'] = references
    if extracted_citations:
        result['citations'] = extracted_citations
    elif 'citations' in config:
        result['citations'] = config['citations']
        
    return result

def generate_user_data_from_file(docx_path: str = None, config_path: Optional[str] = None, parsed_data: dict = None) -> Dict[str, Any]:
    # 统一加载 config
    if config_path is None:
        config_path = Path(__file__).parent / 'user_config.json'
    config = load_config(str(config_path))
    
    # 如果没有传入已解析的数据，则调用 parse_full_docx 进行解析
    if not parsed_data:
        from full_style_docx_fixer.utils.parse_full_docx import parse_full_docx
        parsed_data = parse_full_docx(docx_path)
    
    # 从字典中安全地取出这两部分数据
    docx_infos = parsed_data.get("docx_infos", [])
    citations = parsed_data.get("citations",[])
    
    # 【修复】调用生成函数时，把 citations 也传进去，并且记得 return！
    return generate_user_data(docx_infos, config, extracted_citations=citations)
    


def save_user_data(data: Dict[str, Any], output_path: str):
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1:
        docx_path = sys.argv[1]
    else:
        docx_path = 'data/full_template_v7.docx'
    
    if len(sys.argv) > 2:
        output_path = sys.argv[2]
    else:
        output_path = 'data/generated_user_data.json'
    
    result = generate_user_data_from_file(docx_path)
    save_user_data(result, output_path)
    
    print(f'转换完成，输出文件: {output_path}')
