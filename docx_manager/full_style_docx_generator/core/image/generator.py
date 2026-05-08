import re
import json
from typing import List, Optional, Tuple
from .models import ImageDescription, ImageGroup, ImageGroupingResponse, HeadingSectionResponse, ImageGroupPlacement, CaptionListResponse

_FIG_REF_RE = re.compile(r'图\s*[\d一二三四五六七八九十]|如图|图示|下图|上图|见图|参见图')
_FRONT_MATTER_RE = re.compile(r'^(abstract|目\s*录|摘\s*要|contents)$', re.IGNORECASE)
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


def _select_section_by_heading(
    image_list: list,
    headings: list,
    user_instruction: str = "",
) -> Optional[str]:
    """
    第一跳：让 LLM 从标题列表中选出最适合放置该图片组的章节，返回 heading id。
    失败时返回 None，由调用方降级处理。
    图片编号在 image_list 中为 1-based（image_number 字段）。
    """
    from utils.base_agent import call_structured

    instruction_hint = (
        f"\n【用户说明】{user_instruction.strip()}\n"
        "图片编号从1开始，请根据用户说明优先确定图片所属章节，再结合图片内容判断。"
        if user_instruction and user_instruction.strip()
        else ""
    )

    system_prompt = (
        "你是一个学术文档图片定位助手。根据图片描述，从文档标题列表中选出最适合放置该图片组的章节。\n"
        "选择标准：优先遵循用户说明；其次看图片内容与章节主题是否相关。\n"
        "输出选中标题的 heading_id（即标题的 id 字段）。"
        + instruction_hint
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
    heading_pos = next(
        (i for i, e in enumerate(structured_elements) if e.get('id') == heading_id),
        None,
    )
    if heading_pos is None:
        return [{"index": i, "content": p.get("content", "")[:_MAX_PARA_LEN]} for i, p in enumerate(paragraphs)]

    next_heading_pos = next(
        (i for i in range(heading_pos + 1, len(structured_elements))
         if structured_elements[i].get('type', '') in _HEADING_TYPES),
        len(structured_elements),
    )

    section_ids = {e['id'] for e in structured_elements[heading_pos:next_heading_pos]}

    result = [
        {"index": i, "content": p.get("content", "")[:_MAX_PARA_LEN]}
        for i, p in enumerate(paragraphs)
        if p.get('id') in section_ids
    ]

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


def _strip_caption_prefixes(descriptions: List[ImageDescription]) -> None:
    """
    批量清洗 suggested_caption：剥除任何图号前缀（图1-1/图  /Figure 1 等），
    原地更新为纯描述文字，保证后续 _assign_figure_numbers 可直接拼接。
    """
    from utils.base_agent import call_structured

    raw = [d.suggested_caption for d in descriptions]
    system_prompt = (
        "你是一个文本清洗助手。将每条图题去掉开头的图号前缀，只保留描述内容。\n"
        "前缀形式多样，如：'图1-1 '、'图  '、'图：'、'Figure 1: ' 等，全部去掉。\n"
        "输出 captions 列表，与输入等长，顺序一致，每项只含纯描述文字。"
    )
    user_prompt = json.dumps(raw, ensure_ascii=False)
    response = call_structured(system_prompt, user_prompt, CaptionListResponse)

    for i, desc in enumerate(descriptions):
        if i < len(response.captions):
            desc.suggested_caption = response.captions[i].strip()


def _group_images(descriptions: List[ImageDescription], user_instruction: str = "") -> List[List[int]]:
    """
    Stage 2：纯文本模型根据描述分组。
    图片以 1-based 编号展示给 LLM（image_number 字段），输出后转回 0-based。
    """
    from utils.base_agent import call_structured

    desc_list = [
        {
            "image_number": i + 1,  # 1-based，与用户说的"第一张图"对齐
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
        "图片编号从1开始（image_number 字段）。\n"
        "分组原则：同一实验/同一方法/同一章节的图片归为一组；不相关的各自单独成组。\n"
        "每张图片只能出现在一个组中，所有图片必须被分配。\n"
        "输出 groups 列表，每个子列表填写该组图片的 image_number（1-based）。"
        + grouping_note
    )
    user_prompt = json.dumps(desc_list, ensure_ascii=False)

    response = call_structured(system_prompt, user_prompt, ImageGroupingResponse)
    # 转回 0-based
    return [[num - 1 for num in group] for group in response.groups]


def _place_group(
    image_indices: List[int],
    descriptions: List[ImageDescription],
    paragraphs: list,
    instruction_note: str = "",
    structured_elements: Optional[list] = None,
    user_instruction: str = "",
) -> ImageGroup:
    """
    Stage 3：确定单个图片组的插入位置。
    LLM 只输出 anchor_idx，image_indices 和 captions 直接从已有数据填充。
    图片以 1-based image_number 展示给 LLM。
    """
    from utils.base_agent import call_structured

    # 图片信息用 1-based image_number 展示，避免与用户"第X张图"错位
    image_list = [
        {
            "image_number": idx + 1,
            "main_content": descriptions[idx].main_content,
            "key_concepts": descriptions[idx].key_concepts,
        }
        for idx in image_indices
    ]

    all_concepts = [kw for idx in image_indices for kw in descriptions[idx].key_concepts]
    candidate_paragraphs, is_fallback = _filter_candidate_paragraphs(paragraphs, all_concepts)

    if is_fallback and structured_elements:
        # 只发 id + content，去掉 type 等冗余字段
        headings = [
            {"id": e["id"], "content": e.get("content", "")[:80]}
            for e in structured_elements
            if e.get("type", "") in _HEADING_TYPES
        ]
        if headings:
            print(f"[image.generator] 关键词未命中，启动两阶段定位（{len(headings)} 个标题）")
            heading_id = _select_section_by_heading(image_list, headings, user_instruction)
            if heading_id:
                candidate_paragraphs = _get_section_paragraphs(heading_id, paragraphs, structured_elements)
                print(f"[image.generator] 选中标题 {heading_id}，章节内候选段落 {len(candidate_paragraphs)} 个")

    system_prompt = (
        "你是一个学术文档图片定位助手。根据图片描述和候选段落，确定这组图片的插入位置。\n\n"
        "从候选段落中选一个段落，图片组将插入在该段落之后。\n"
        "输出该段落的 index 值（anchor_idx）。\n"
        "优先找引用 key_concepts 关键词或描述图片内容的段落。"
        + instruction_note
    )
    user_prompt = json.dumps(
        {"images": image_list, "paragraphs": candidate_paragraphs},
        ensure_ascii=False,
    )

    placement = call_structured(system_prompt, user_prompt, ImageGroupPlacement)

    return ImageGroup(
        image_indices=image_indices,
        anchor_idx=placement.anchor_idx,
        captions=[descriptions[idx].suggested_caption for idx in image_indices],
    )


def _get_body_heading1s(structured_elements: list) -> list:
    """返回正文区域的 heading1 列表（跳过 Abstract/目录/摘要 等前言性标题）。"""
    all_h1 = [e for e in structured_elements if e.get('type') == 'heading1']
    last_front = -1
    for i, h in enumerate(all_h1):
        if _FRONT_MATTER_RE.match(h.get('content', '').strip()):
            last_front = i
    return all_h1[last_front + 1:]


def _assign_figure_numbers(
    result: List['ImageGroup'],
    paragraphs: list,
    structured_elements: list,
) -> None:
    """
    原地替换每个 ImageGroup.captions 的图号前缀（图\\s* → 图X-Y ）。
    章号 = 正文第 N 个 heading1（N 从 1 起）。
    序号 = 同章内按 anchor_idx 升序排列后的出现顺序（从 1 起）。
    """
    se_id_to_pos = {e['id']: i for i, e in enumerate(structured_elements) if 'id' in e}
    body_h1 = _get_body_heading1s(structured_elements)
    body_h1_id_to_chapter = {h['id']: idx + 1 for idx, h in enumerate(body_h1)}
    body_h1_ids_set = set(body_h1_id_to_chapter)

    def _chapter_of(anchor_idx: int) -> Optional[int]:
        para = paragraphs[anchor_idx] if anchor_idx < len(paragraphs) else None
        if para is None:
            return None
        se_pos = se_id_to_pos.get(para.get('id'))
        if se_pos is None:
            return None
        for i in range(se_pos, -1, -1):
            eid = structured_elements[i].get('id')
            if eid in body_h1_ids_set:
                return body_h1_id_to_chapter[eid]
        return None

    # 按文档顺序（anchor_idx 升序）分配章内序号
    ordered = sorted(range(len(result)), key=lambda i: result[i].anchor_idx)
    chapter_counter: dict[int, int] = {}
    for i in ordered:
        group = result[i]
        chapter = _chapter_of(group.anchor_idx)
        if chapter is None:
            continue
        new_captions = []
        for cap in group.captions:
            chapter_counter[chapter] = chapter_counter.get(chapter, 0) + 1
            fig_num = chapter_counter[chapter]
            new_captions.append(f"图{chapter}-{fig_num} {cap}")
        group.captions = new_captions


def generate(
    images: list,
    paragraphs: list,
    user_instruction: str = "",
    structured_elements: Optional[list] = None,
) -> List[ImageGroup]:
    instruction_note = (
        f"\n\n【用户补充说明】\n{user_instruction.strip()}\n"
        "图片编号从1开始，请优先遵循以上说明确定图片位置。"
        if user_instruction and user_instruction.strip()
        else ""
    )

    # Stage 1: 多模态描述生成
    descriptions = [_describe_image(img) for img in images]
    print(f"[image.generator] Stage 1 完成：{len(descriptions)} 张图片描述已生成")

    # Stage 1.5: 清洗图题前缀，保证 suggested_caption 为纯描述
    _strip_caption_prefixes(descriptions)
    print(f"[image.generator] Stage 1.5 完成：图题前缀已清洗")

    # Stage 2: 纯文本分组（1-based 展示，内部转回 0-based）
    groups = _group_images(descriptions, user_instruction)
    print(f"[image.generator] Stage 2 完成：分为 {len(groups)} 组")

    # Stage 3: 逐组定位（只让 LLM 输出 anchor_idx）
    result = []
    for g in groups:
        try:
            result.append(_place_group(g, descriptions, paragraphs, instruction_note, structured_elements, user_instruction))
        except Exception as e:
            print(f"[image.generator] 图片组 {g} 定位失败，跳过: {e}")
    print(f"[image.generator] Stage 3 完成：{len(result)}/{len(groups)} 组定位成功")

    if structured_elements:
        _assign_figure_numbers(result, paragraphs, structured_elements)
        print(f"[image.generator] 图号分配完成")

    return result
