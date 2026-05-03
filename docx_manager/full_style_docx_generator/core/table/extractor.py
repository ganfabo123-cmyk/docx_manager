from utils.base_agent import call_structured
from .models import TableExtractResponse

_SYSTEM_PROMPT = """\
## 角色
你是一个文档结构化处理助手，专门负责从非规范文本中提取表格数据。

## 任务背景
用户文档由大模型生成，其中的表格以 Markdown 竖线语法书写，但格式可能不规范：
列宽对齐混乱、存在分隔行（|---|）、单元格内含加粗/斜体标记、部分行缺少首尾竖线等。
这些内容在解析后以逐行文本的形式输入给你，你需要将其还原为结构化表格。

## 任务
判断输入内容是否为表格，若是则提取为标准二维数组；若不是则标记后跳过。

## 执行步骤
1. 判断：输入是否描述了一个结构化表格（含列标题 + 数据行）？
   - 若否（如纯文字、公式、代码片段），将 is_not_table 置为 True，content 填 []，title 填 ""，结束。
2. 提取标题：检查表格正文之前是否有独立标题行（如"表1 xxx"），有则提取到 title；否则 title 填 ""。
3. 提取列标题：将表头行的各列文字作为 content[0]，保留单位信息（如"温度(℃)"）。
4. 提取数据行：每条记录对应 content 中的一行，顺序与原文一致。
5. 清洗：去除所有 Markdown 格式符号（** * ` ~~ 等），保留原始数值和文字。
"""


def extract_table(raw_text: str, existing_title: str | None = None) -> list[dict] | None:
    """
    Input:  raw_text       — 同一张表格的所有行拼合后的原始文本
            existing_title — 若前置元素已是表题，传入其文本；否则为 None，由 LLM 生成
    Output: [body_block, table_block]（无 existing_title）
            或 [table_block]（有 existing_title，body 由原文档元素承担）
            若判断非表格则返回 None
    """
    if existing_title is not None:
        title_instruction = f'- title: str（直接使用已知表题"{existing_title}"，不要修改）'
    else:
        title_instruction = "- title: str（根据表格内容起一个简短标题，无法判断则填空字符串）"

    user_prompt = (
        f"请处理以下输入内容，按要求填写响应字段。\n\n"
        + (f"【已知表题】{existing_title}\n\n" if existing_title is not None else "")
        + f"【输入内容】\n{raw_text}\n\n"
        f"【格式要求】\n"
        f"- is_not_table: bool\n"
        f"{title_instruction}\n"
        f"- content: list[list[str]]（二维数组，非表格填 []）"
    )

    result: TableExtractResponse = call_structured(_SYSTEM_PROMPT, user_prompt, TableExtractResponse)

    if result.is_not_table:
        return None

    rows = len(result.content)
    cols = max((len(row) for row in result.content), default=0)

    blocks = []
    if existing_title is None:
        blocks.append({
            "type":    "body",
            "content": result.title,
            "style":   {"style_name": "Normal"},
        })
    blocks.append({
        "type":    "table",
        "content": result.content,
        "style":   {"style_name": "Table", "rows": rows, "cols": cols},
    })
    return blocks
