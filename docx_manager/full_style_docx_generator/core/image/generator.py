import re
import json
from typing import List
from .models import ImageGroup, ImageGroupListResponse

# 匹配学术文档中常见的图片引用模式
_FIG_REF_RE = re.compile(r'图\s*[\d一二三四五六七八九十]|如图|图示|下图|上图|见图|参见图')

_MAX_PARA_LEN = 200  # 单段落发给 LLM 的最大字符数


def _filter_candidate_paragraphs(paragraphs: list, window: int = 1) -> list:
    """
    返回含图片引用关键词的段落及前后 window 个上下文段落。
    保留原始 index，使 LLM 输出的 anchor_idx 可直接用于 backfill 查找。
    若无任何命中，退为返回全部段落（降级兜底）。
    """
    n = len(paragraphs)
    candidate_indices: set[int] = set()

    for i, p in enumerate(paragraphs):
        if _FIG_REF_RE.search(p.get('content', '')):
            for j in range(max(0, i - window), min(n, i + window + 1)):
                candidate_indices.add(j)

    if not candidate_indices:
        # 没有关键词命中，发全部（截断每段长度）
        return [
            {"index": i, "content": p.get("content", "")[:_MAX_PARA_LEN]}
            for i, p in enumerate(paragraphs)
        ]

    return [
        {
            "index": i,
            "content": paragraphs[i].get("content", "")[:_MAX_PARA_LEN],
        }
        for i in sorted(candidate_indices)
    ]


def generate(images: list, paragraphs: list, user_instruction: str = "") -> List[ImageGroup]:
    from utils.base_agent import call_structured

    image_list = [
        {"index": i, "caption": img.get("caption", "")}
        for i, img in enumerate(images)
    ]
    candidate_paragraphs = _filter_candidate_paragraphs(paragraphs)

    instruction_block = (
        f"\n\n【用户补充说明】\n{user_instruction.strip()}\n请优先遵循以上说明确定图片位置。"
        if user_instruction and user_instruction.strip()
        else ""
    )

    system_prompt = (
        "你是一个学术文档图片定位助手。给定一组图片信息和文档段落（段落已预筛选，"
        "index 为该段落在完整文档中的原始编号），完成以下任务：\n\n"
        "1. 将相关联的图片分为同一组（同一实验不同视角、同一章节引用的图片等）\n"
        "2. 确定每组图片的展示顺序（image_indices 按顺序排列）\n"
        "3. 确定每组图片的 anchor_idx：填写段落的原始 index 值，图片插入该段落之后\n"
        "   规则：优先选明确引用或描述该组图片的段落；若无明确引用，选语义最相关的段落\n"
        "4. 为每张图片生成图题（captions，与 image_indices 等长）：\n"
        "   - 若图片已有非空 caption，直接使用原 caption\n"
        "   - 若 caption 为空，根据锚点段落内容推断图题，格式为 '图 X  简短描述'\n\n"
        "输出 groups 列表，每个 group 包含 image_indices、anchor_idx、captions。"
        + instruction_block
    )

    user_prompt = json.dumps(
        {"images": image_list, "paragraphs": candidate_paragraphs},
        ensure_ascii=False,
    )

    response = call_structured(system_prompt, user_prompt, ImageGroupListResponse)
    return response.groups
