import json
from typing import List, Dict, Any
from .models import ShortBlockListResponse


SYSTEM_PROMPT = (
    "你是一个文档结构分析器。你会收到一组短文本块（每个含 id 和 content），"
    "任务是判断每个文本块是正文还是标题，以及标题的层级。\n\n"
    "分类规则：\n"
    "- heading1：一级标题（如 '第一章 绪论'、'1 引言'、'Abstract'、'摘要' 等章节顶层标题）\n"
    "- heading2：二级标题（如 '1.1 研究背景'、'2.1 相关工作' 等）\n"
    "- heading3：三级标题（如 '1.1.1 具体内容' 等）\n"
    "- normal：普通正文、无意义短文本或无法确定层级的内容\n\n"
    "只输出每个元素的 id 和 type，不要解释。"
)

BATCH_SIZE = 10


def classify_short_blocks(short_blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    LLM-classify each short block as heading1/heading2/heading3/normal.
    Returns list of {id, content, type}.
    """
    from utils.base_agent import call_structured

    if not short_blocks:
        return []

    results = []
    total_batches = (len(short_blocks) + BATCH_SIZE - 1) // BATCH_SIZE

    for i in range(0, len(short_blocks), BATCH_SIZE):
        batch = short_blocks[i:i + BATCH_SIZE]
        user_prompt = json.dumps(batch, ensure_ascii=False)
        response = call_structured(SYSTEM_PROMPT, user_prompt, ShortBlockListResponse)

        id_to_type = {item.id: item.type for item in response.items}
        for block in batch:
            results.append({
                "id": block["id"],
                "content": block["content"],
                "type": id_to_type.get(block["id"], "normal"),
            })

        print(f"🤖 [LLM] 批次 {i // BATCH_SIZE + 1}/{total_batches}：分类 {len(batch)} 个短文本块")

    return results
