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
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
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

def generate_user_data(docx_info: List[Dict[str, Any]], config: Dict[str, Any]) -> Dict[str, Any]:
    content = []
    toc_entries = []
    references = []
    
    i = 0
    while i < len(docx_info):
        item = docx_info[i]
        item_type = item.get('type', '')
        
        if item_type.startswith('toc'):
            toc_entry = extract_toc_entry(item)
            if toc_entry:
                toc_entries.append(toc_entry)
            i += 1
            continue
        
        if is_section_type(item_type):
            content.append(convert_section(item, config))
            i += 1
        elif item_type.startswith('heading'):
            heading_title = item.get('value', '')
            section_type = is_special_section_title(heading_title)
            
            if section_type:
                section_content = []
                j = i + 1
                while j < len(docx_info):
                    next_item = docx_info[j]
                    next_type = next_item.get('type', '')
                    if next_type.startswith('heading'):
                        break
                    if next_type == 'body':
                        section_content.append(next_item.get('value', ''))
                    j += 1
                
                combined_content = '\n\n'.join(section_content)
                
                if section_type == 'references':
                    ref_pattern = re.compile(r'\[(\d+)\](.+)')
                    for body_text in section_content:
                        match = ref_pattern.match(body_text.strip())
                        if match:
                            ref_id = int(match.group(1))
                            ref_text = match.group(2).strip()
                            references.append({"id": ref_id, "text": ref_text})
                else:
                    section_item = {
                        'type': 'section',
                        'section_type': section_type,
                        'toc_exclude': True if section_type in ['abstract', 'abstract_en', 'custom'] else False,
                        'value': combined_content
                    }
                    if section_type == 'custom':
                        section_item['title'] = heading_title
                    content.append(section_item)
                
                i = j
            else:
                content.append(convert_heading(item, config))
                i += 1
        elif item_type == 'body':
            content.append(convert_body(item))
            i += 1
        elif item_type == 'table':
            content.append(convert_table(item))
            i += 1
        elif item_type == 'image':
            content.append(convert_image(item, config))
            i += 1
        elif item_type == 'formula':
            content.append(convert_formula(item, config))
            i += 1
        elif item_type == 'reference':
            content.append(convert_reference(item))
            i += 1
        else:
            i += 1
    
    result = {
        '_doc': '由 parse_full_docx 生成的数据转换而来',
        '_tips': {
            '图片_path模式': '"path": "/absolute/path/to/image.png"',
            '图片_base64模式': '"base64": "<base64字符串>", "ext": "png"',
            'toc_exclude': 'true → 标题不进 TOC 域（但可手动写入 toc_entries）',
            '公式_omml': '直接嵌入 Office Open Math XML，零依赖',
            '公式_latex': '需要 pip install latex2mathml，否则退化为纯文本'
        },
        'page_footer_config': config.get('page_footer_config', []),
        'toc_mode': config.get('toc_mode', 'manual'),
        'toc_entries': toc_entries,
        'content': content
    }
    
    if references:
        result['references'] = references
    
    if 'citations' in config:
        result['citations'] = config['citations']
    
    return result


def generate_user_data_from_file(docx_path: str, config_path: Optional[str] = None) -> Dict[str, Any]:
    from llm_data_collector.utils.parse_full_docx import parse_full_docx
    
    docx_info = parse_full_docx(docx_path)
    
    if config_path is None:
        config_path = Path(__file__).parent / 'user_config.json'
    
    config = load_config(str(config_path))
    
    return generate_user_data(docx_info.get("docx_infos"), config)


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
