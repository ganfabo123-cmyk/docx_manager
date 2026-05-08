import re
import json
from typing import List, Optional, Tuple
from .models import ImageDescription, ImageGroup, ImageGroupingResponse, HeadingSectionResponse

_FIG_REF_RE = re.compile(r'图\s*[\d一二三四五六七八九十]|如图|图示|下图|上图|见图|参见图')
_MAX_PARA_LEN = 200
_HEADING_TYPES = {'heading1', 'heading2', 'heading3', 'heading', 'title'}


def _filter_candidate_paragraphs(
    paragraphs: list,
    extra_keywords: List[str] = None,
    window: int = 1,
) -> Tuple[list, bool]:
    """
    返回含图片引用关键词的段落及前后 window 个上下文段落，同时返回是否触发了兜底。
    保留原始 index，使 LLM 输出的 anchor_idx 可直接用于 backfill 查找。
    """
    n = len(paragraphs)
    candidate_indices: set[int] = set()

    patterns = [_FIG_REF_RE.pattern]
    if extra_keywords:
        escaped = [re.escape(kw) for kw in extra_keywords if len(kw) >= 2]
        if escaped:
            patterns.extend(escaped)
    combined_re = re.compile('|'.join(patterns), re.IGNORECASE)

    for i, p in enumerate(paragraphs):
        if combined_re.search(p.get('content', '')):
            for j in range(max(0, i - window), min(n, i + window + 1)):
                candidate_indices.add(j)

    if not candidate_indices:
        return [
            {"index": i, "content": p.get("content", "")[:_MAX_PARA_LEN]}
            for i, p in enumerate(paragraphs)
        ], True  # is_fallback=True

    return [
        {"index": i, "content": paragraphs[i].get("content", "")[:_MAX_PARA_LEN]}
        for i in sorted(candidate_indices)
    ], False


def _select_section_by_heading(image_list: list, headings: list) -> Optional[str]:
    """
    第一跳：让 LLM 从标题列表中选出最适合放置该图片组的章节，返回 heading id。
    失败时返回 None，由调用方降级处理。
    """
    from utils.base_agent import call_structured

    system_prompt = (
        "你是一个学术文档图片定位助手。根据图片描述，从文档标题列表中选出最适合放置该图片组的章节。\n"
        "选择标准：图片内容与该章节主题最相关，图片应插入在该章节的正文中。\n"
        "输出选中标题的 heading_id（即标题的 id 字段）。"
    )
    user_prompt = json.dumps(
        {"images": image_list, "headings": headings},
        ensure_ascii=False,
    )

    try:
        response = call_structured(system_prompt, user_prompt, HeadingSectionResponse)
        return response.heading_id
    except Exception as e:
        print(f"[image.generator] 标题选择失败，降级为全文搜索: {e}")
        return None


def _get_section_paragraphs(
    heading_id: str,
    paragraphs: list,
    structured_elements: list,
) -> list:
    """
    第二跳的候选段落：从 structured_elements 中定位 heading_id 所在位置，
    取到下一个标题之前的所有元素，再从 paragraphs 中筛出对应的段落（保留原始 index）。
    """
    # 找到标题在 structured_elements 中的位置
    heading_pos = next(
        (i for i, e in enumerate(structured_elements) if e.get('id') == heading_id),
        None,
    )
    if heading_pos is None:
        return [{"index": i, "content": p.get("content", "")[:_MAX_PARA_LEN]} for i, p in enumerate(paragraphs)]

    # 找到下一个标题位置
    next_heading_pos = next(
        (i for i in range(heading_pos + 1, len(structured_elements))
         if structured_elements[i].get('type', '') in _HEADING_TYPES),
        len(structured_elements),
    )

    # 该章节内的元素 id 集合（含标题本身，方便锚点落在标题上）
    section_ids = {e['id'] for e in structured_elements[heading_pos:next_heading_pos]}

    # 从 paragraphs 中筛出属于该章节的段落，保留原始 index
    result = [
        {"index": i, "content": p.get("content", "")[:_MAX_PARA_LEN]}
        for i, p in enumerate(paragraphs)
        if p.get('id') in section_ids
    ]

    # 若章节内没有匹配到任何段落（数据不一致），降级返回全文
    if not result:
        return [{"index": i, "content": p.get("content", "")[:_MAX_PARA_LEN]} for i, p in enumerate(paragraphs)]

    return result


def _describe_image(image: dict) -> ImageDescription:
    """Stage 1：多模态模型分析单张图片，生成结构化描述。"""
    from utils.base_agent import call_structured_with_image

    system_prompt = (
        "你是一个学术图片分析助手，请仔细分析图片内容并按字段输出结构化描述。"
        "字段说明：\n"
        "- figure_type：图片类型（折线图/柱状图/框架图/示意图/实验截图/表格截图/其他）\n"
        "- topic_summary：一句话主题摘要\n"
        "- main_content：2-4句完整内容描述\n"
        "- key_concepts：图中出现的技术关键词列表\n"
        "- suggested_caption：推荐图题，格式为'图  简短描述'（图号留空）"
    )
    user_prompt = "请分析这张图片，填写各字段。"
    return call_structured_with_image(system_prompt, user_prompt, image.get('base64', ''), ImageDescription)


def _group_images(descriptions: List[ImageDescription], user_instruction: str = "") -> List[List[int]]:
    """Stage 2：纯文本模型根据描述分组，不需要看图片。"""
    from utils.base_agent import call_structured

    desc_list = [
        {
            "index": i,
            "figure_type": d.figure_type,
            "topic_summary": d.topic_summary,
        }
        for i, d in enumerate(descriptions)
    ]

    grouping_note = (
        f"\n\n【用户补充说明】\n{user_instruction.strip()}\n"
        "请从以上说明中提取与图片分组相关的指令（如哪些图应合并、哪些图应拆分等），"
        "优先遵循这些分组要求；与分组无关的内容忽略。"
        if user_instruction and user_instruction.strip()
        else ""
    )

    system_prompt = (
        "你是一个文档图片分组助手。根据图片类型和主题摘要，将相关联的图片分为同一组。\n"
        "分组原则：同一实验/同一方法/同一章节的图片归为一组；不相关的各自单独成组。\n"
        "每张图片只能出现在一个组中，所有图片必须被分配。\n"
        "输出 groups 列表，每个元素是该组图片的 index 列表。"
        + grouping_note
    )
    user_prompt = json.dumps(desc_list, ensure_ascii=False)

    response = call_structured(system_prompt, user_prompt, ImageGroupingResponse)
    return response.groups


def _place_group(
    image_indices: List[int],
    descriptions: List[ImageDescription],
    paragraphs: list,
    instruction_note: str = "",
    structured_elements: Optional[list] = None,
) -> ImageGroup:
    """Stage 3：纯文本模型为单个图片组确定锚点段落和图题。"""
    from utils.base_agent import call_structured

    image_list = [
        {
            "index": idx,
            "main_content": descriptions[idx].main_content,
            "key_concepts": descriptions[idx].key_concepts,
            "suggested_caption": descriptions[idx].suggested_caption,
        }
        for idx in image_indices
    ]

    all_concepts = [kw for idx in image_indices for kw in descriptions[idx].key_concepts]
    candidate_paragraphs, is_fallback = _filter_candidate_paragraphs(paragraphs, all_concepts)

    if is_fallback and structured_elements:
        headings = [
            {"id": e["id"], "type": e.get("type", ""), "content": e.get("content", "")[:80]}
            for e in structured_elements
            if e.get("type", "") in _HEADING_TYPES
        ]
        if headings:
            print(f"[image.generator] 关键词未命中，启动两阶段定位（{len(headings)} 个标题）")
            heading_id = _select_section_by_heading(image_list, headings)
            if heading_id:
                candidate_paragraphs = _get_section_paragraphs(heading_id, paragraphs, structured_elements)
                print(f"[image.generator] 选中标题 {heading_id}，章节内候选段落 {len(candidate_paragraphs)} 个")

    system_prompt = (
        "你是一个学术文档图片定位助手。根据图片描述和候选段落，确定这组图片的插入位置和图题。\n\n"
        "1. anchor_idx：图片组插入在该段落之后，填写段落的 index 值\n"
        "   优先找引用 key_concepts 关键词或描述图片内容的段落\n"
        "2. captions：每张图片一个图题，优先使用 suggested_caption，可根据上下文微调\n"
        "3. image_indices：保持传入顺序不变"
        + instruction_note
    )
    user_prompt = json.dumps(
        {"images": image_list, "paragraphs": candidate_paragraphs},
        ensure_ascii=False,
    )

    return call_structured(system_prompt, user_prompt, ImageGroup)


def generate(
    images: list,
    paragraphs: list,
    user_instruction: str = "",
    structured_elements: Optional[list] = None,
) -> List[ImageGroup]:
    instruction_note = (
        f"\n\n【用户补充说明】\n{user_instruction.strip()}\n请优先遵循以上说明确定图片位置。"
        if user_instruction and user_instruction.strip()
        else ""
    )

    # Stage 1: 多模态描述生成
    descriptions = [_describe_image(img) for img in images]
    print(f"[image.generator] Stage 1 完成：{len(descriptions)} 张图片描述已生成")

    # Stage 2: 纯文本分组
    groups = _group_images(descriptions, user_instruction)
    print(f"[image.generator] Stage 2 完成：分为 {len(groups)} 组")

    # Stage 3: 逐组定位
    result = [
        _place_group(g, descriptions, paragraphs, instruction_note, structured_elements)
        for g in groups
    ]
    print(f"[image.generator] Stage 3 完成：所有组定位完毕")

    return result
