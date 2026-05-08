import re
import json
from typing import List
from .models import ImageDescription, ImageGroup, ImageGroupingResponse

_FIG_REF_RE = re.compile(r'图\s*[\d一二三四五六七八九十]|如图|图示|下图|上图|见图|参见图')
_MAX_PARA_LEN = 200


def _filter_candidate_paragraphs(
    paragraphs: list,
    extra_keywords: List[str] = None,
    window: int = 1,
) -> list:
    """
    返回含图片引用关键词（或 extra_keywords）的段落及前后 window 个上下文段落。
    保留原始 index，使 LLM 输出的 anchor_idx 可直接用于 backfill 查找。
    若无任何命中，退为返回全部段落（降级兜底）。
    """
    n = len(paragraphs)
    candidate_indices: set[int] = set()

    patterns = [_FIG_REF_RE.pattern]
    if extra_keywords:
        escaped = [re.escape(kw) for kw in extra_keywords if len(kw) >= 2]
        if escaped:
            patterns.extend(escaped)
    combined_re = re.compile('|'.join(patterns))

    for i, p in enumerate(paragraphs):
        if combined_re.search(p.get('content', '')):
            for j in range(max(0, i - window), min(n, i + window + 1)):
                candidate_indices.add(j)

    if not candidate_indices:
        return [
            {"index": i, "content": p.get("content", "")[:_MAX_PARA_LEN]}
            for i, p in enumerate(paragraphs)
        ]

    return [
        {"index": i, "content": paragraphs[i].get("content", "")[:_MAX_PARA_LEN]}
        for i in sorted(candidate_indices)
    ]


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
    candidate_paragraphs = _filter_candidate_paragraphs(paragraphs, all_concepts)

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


def generate(images: list, paragraphs: list, user_instruction: str = "") -> List[ImageGroup]:
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
    result = [_place_group(g, descriptions, paragraphs, instruction_note) for g in groups]
    print(f"[image.generator] Stage 3 完成：所有组定位完毕")

    return result
