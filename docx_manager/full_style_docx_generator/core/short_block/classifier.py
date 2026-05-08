import json
from typing import List, Dict, Any
from .models import ShortBlockListResponse, HeadingCorrectionResponse


SYSTEM_PROMPT = (
    "你是一个文档结构分析器。你会收到一组短文本块（每个含 id 和 content），"
    "任务是判断每个文本块是正文还是标题，以及标题的层级。\n\n"
    "分类规则：\n"
    "- heading1：一级标题，即文档最顶层的章节标题。包括：\n"
    "  * 以中文数字加顿号开头，如 '一、实验内容介绍'、'二、核心算法'、'三、结论'\n"
    "  * 以阿拉伯数字（不含小数点）开头，如 '1 引言'、'2 相关工作'\n"
    "  * 以章序词开头，如 '第一章 绪论'、'第二章 方法'\n"
    "  * 固定前言标题，如 'Abstract'、'摘要'、'摘  要'、'结论'、'参考文献'\n"
    "- heading2：二级标题，即章节内的小节标题。包括：\n"
    "  * 以 'x.y' 格式编号，如 '1.1 研究背景'、'2.1 相关工作'\n"
    "  * 无编号但明显是小节的短语，如 '实验背景与动机'、'模型架构设计'、'结果分析'\n"
    "- heading3：三级标题，如 '1.1.1 具体内容' 等\n"
    "- normal：普通正文、无意义短文本或无法确定层级的内容\n\n"
    "关键判断规则：\n"
    "- 以'一、'/'二、'/'三、'等中文数字加顿号开头的，一定是 heading1\n"
    "- 含小数点编号（如 1.1、2.3）的，一定是 heading2 或 heading3\n"
    "- 正文段落通常超过 20 个字，标题通常很短\n\n"
    "只输出每个元素的 id 和 type，不要解释。"
)

NORMALIZER_PROMPT = (
    "你是一个标题层级校正器。以下是文档中所有标题，按出现顺序排列，已有初步分类。\n"
    "请根据标题内容和整体结构关系，输出每个标题修正后的层级。\n\n"
    "规则（优先级从高到低）：\n"
    "1. 以'一、二、三、四、五、六、七、八、九、十'等中文数字加顿号开头 → 必须是 heading1\n"
    "2. 以'第X章'格式开头 → 必须是 heading1\n"
    "3. 固定名称（摘要 / 摘  要 / Abstract / 结论 / 参考文献 / 目录）→ heading1\n"
    "4. 编号形如 'x.y'（含小数点，如 1.1、2.3）→ heading2\n"
    "5. 编号形如 'x.y.z' → heading3\n"
    "6. 在已有'一、二、三'章节编号的文档中，其余无编号短标题一律是 heading2，不可升为 heading1\n\n"
    "只输出每个标题的 id 和 corrected_type，不要解释。"
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


def normalize_heading_structure(classified_results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Single-pass LLM call over all headings to correct hierarchy.
    Non-heading items pass through unchanged.
    """
    from utils.base_agent import call_structured

    headings = [r for r in classified_results if r.get("type", "").startswith("heading")]
    if not headings:
        return classified_results

    payload = [{"id": h["id"], "content": h["content"], "type": h["type"]} for h in headings]
    response = call_structured(NORMALIZER_PROMPT, json.dumps(payload, ensure_ascii=False), HeadingCorrectionResponse)

    corrections = {item.id: item.corrected_type for item in response.items}

    result = []
    for r in classified_results:
        if r["id"] in corrections:
            result.append({**r, "type": corrections[r["id"]]})
        else:
            result.append(r)

    corrected_count = sum(
        1 for r in classified_results
        if r["id"] in corrections and corrections[r["id"]] != r.get("type")
    )
    print(f"🤖 [LLM-NORM] 标题结构校正完成：{len(headings)} 个标题，修正 {corrected_count} 个")
    return result
