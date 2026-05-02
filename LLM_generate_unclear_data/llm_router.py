import json
import os
import re

from base_agent import BaseAgent
from formula_generate import BlockFormula, FormulaItem, InlineFormula

_agent = BaseAgent(
    endpoint=os.getenv("LLM_ENDPOINT", ""),
    api_key=os.getenv("LLM_API_KEY", ""),
    model=os.getenv("LLM_MODEL", ""),
)


def _extract_json(text: str):
    """从 LLM 响应中提取 JSON，兼容 ```json ... ``` 包裹格式。"""
    match = re.search(r'```(?:json)?\s*([\s\S]*?)```', text)
    if match:
        return json.loads(match.group(1).strip())
    return json.loads(text.strip())


def route_formula(raw_text: str) -> list[FormulaItem]:
    system_prompt = (
        "你是专业的数学公式识别与转换专家。\n"
        "分析输入文本，识别所有数学公式，以 JSON 数组输出：\n"
        "- 行内公式（嵌在文字中）：{\"type\":\"inline\",\"text_before\":\"前文\",\"formula\":\"LaTeX\",\"text_after\":\"后文\"}\n"
        "- 独立公式（单独成行，结尾可能有编号如 (4-2)）：{\"type\":\"block\",\"label\":\"4-2\",\"formula\":\"LaTeX\"}\n"
        "  无编号则 label 为 \"\"，formula 字段必须是标准 LaTeX。\n"
        "仅输出 JSON 数组，不加任何说明。"
    )
    raw = _agent.chat(raw_text, system_prompt)
    items_data = _extract_json(raw)
    result = []
    for item in items_data:
        if item["type"] == "inline":
            result.append(InlineFormula(
                text_before=item.get("text_before", ""),
                formula=item["formula"],
                text_after=item.get("text_after", ""),
            ))
        else:
            result.append(BlockFormula(
                label=item.get("label", ""),
                formula=item["formula"],
            ))
    return result


def route_table(raw_input: str) -> str:
    system_prompt = (
        "你是专业的数据表格处理专家。\n"
        "将输入的表格数据（混乱文本或已结构化的 JSON 二维数组）整理为以下格式：\n"
        "{\"title\":\"表格标题（无则空字符串）\","
        "\"content\":[[\"列1标题\",\"列2标题\",...],[\"数据\",\"数据\",...],...]}\n"
        "规则：\n"
        "- content 第一行为列标题/属性名，含单位信息\n"
        "- content 其余行为数据行，所有值为字符串\n"
        "- 若输入已是结构化二维数组，提取标题并规范格式\n"
        "仅输出 JSON，不加任何说明。"
    )
    return _agent.chat(raw_input, system_prompt)


def route_confirm(suspected: dict) -> dict:
    """
    Input:  {"formula": [body块], "table": [body块]}  仅含需要 LLM 判断的 body 块
    Output: {"formula": [确认为公式的块], "table": [确认为表格的块]}
    """
    system_prompt = (
        "你是专业的文档内容分析专家。\n"
        "判断以下文本块是否真正包含数学公式或表格数据，返回确认为真的块：\n"
        "{\"formula\":[确认含公式的块],\"table\":[确认含表格的块]}\n"
        "- formula 列表：文本中真正含数学公式（非普通文字描述）\n"
        "- table 列表：文本真正描述了结构化表格数据\n"
        "仅输出 JSON，不加任何说明。"
    )
    raw = _agent.chat(json.dumps(suspected, ensure_ascii=False), system_prompt)
    return _extract_json(raw)


def route_group_images(images: list, paragraphs: list[str]) -> list[list[int]]:
    """
    [待确认] 图片传输格式以平台多模态接口为准，当前以 base64 字段长度描述传入。
    """
    system_prompt = (
        "你是专业的文档图片分析专家。\n"
        "将相关联的图片分组（同一实验/场景/主题的不同视角或步骤归为一组）。\n"
        "参考文档段落辅助判断。\n"
        "输出 JSON 数组，每个子数组为一组图片的索引（从0开始）：[[0,1],[2],[3,4]]\n"
        "仅输出 JSON，不加任何说明。"
    )
    img_desc = "\n".join(
        f"图片{i}: base64长度={len(img.get('base64', '')) if isinstance(img, dict) else len(img)}字节"
        for i, img in enumerate(images)
    )
    para_desc = "文档段落：\n" + "\n".join(f"{i}: {p}" for i, p in enumerate(paragraphs))
    raw = _agent.chat(f"{img_desc}\n{para_desc}", system_prompt)
    return _extract_json(raw)


def route_sort_group_images(image_indices: list[int], paragraphs: list[str]) -> list[int]:
    """
    [待确认] 图片传输格式以平台多模态接口为准。
    """
    system_prompt = (
        "你是专业的文档图片排序专家。\n"
        "根据文档语义和图片内容，确定这组图片的逻辑排列顺序。\n"
        "输出按正确顺序排列的图片索引 JSON 数组：[0,1] 或 [1,0]\n"
        "仅输出 JSON，不加任何说明。"
    )
    para_desc = "文档段落：\n" + "\n".join(f"{i}: {p}" for i, p in enumerate(paragraphs))
    raw = _agent.chat(f"待排序图片索引：{image_indices}\n{para_desc}", system_prompt)
    return _extract_json(raw)


def route_anchor_idx(sorted_groups: list[list[int]], paragraphs: list[str]) -> str:
    system_prompt = (
        "你是专业的文档排版专家。\n"
        "根据文档段落内容和图片组信息，确定每组图片应插入到哪个段落之后。\n"
        "输出 JSON 数组：[{\"image_indices\":[0,1],\"anchor_idx\":2},{\"image_indices\":[2],\"anchor_idx\":5}]\n"
        "anchor_idx 为段落列表的下标（从0开始），图片插入该段落之后。\n"
        "仅输出 JSON，不加任何说明。"
    )
    group_desc = "图片分组（已排序）：\n" + "\n".join(f"组{i}: {g}" for i, g in enumerate(sorted_groups))
    para_desc = "文档段落：\n" + "\n".join(f"{i}: {p}" for i, p in enumerate(paragraphs))
    return _agent.chat(f"{group_desc}\n{para_desc}", system_prompt)
