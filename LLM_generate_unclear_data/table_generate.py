import json
from dataclasses import dataclass


@dataclass
class TableData:
    title: str
    content: list[list[str]]   # 第一行为列标题，其余为数据行（对齐 docx_parser 格式）


def _make_blocks(data: TableData) -> list[dict]:
    """将 TableData 转换为两个标准 JSON 块：body（标题）+ table（数据）。"""
    rows = len(data.content)
    cols = max((len(row) for row in data.content), default=0)
    return [
        {
            "type":    "body",
            "content": data.title,
            "style":   {"style_name": "Normal"},
        },
        {
            "type":    "table",
            "content": data.content,
            "style":   {"style_name": "Table", "rows": rows, "cols": cols},
        },
    ]


def convert(table_json_str: str) -> list[dict]:
    """
    Input:  str — LLM 返回的 JSON，格式：
            {"title": "...", "content": [["列标题",...], ["数据",...], ...]}
    Output: list[dict] — [body_title_block, table_block]，对齐 docx_parser 标准格式
    """
    data = json.loads(table_json_str)
    return _make_blocks(TableData(
        title=data.get("title", ""),
        content=data["content"],
    ))


def generate(blocks: list[dict]) -> list[dict]:
    """
    Input:  list[dict] — 已确认的表格块列表（type=="table" 或 type=="body"，均为表格内容）
    Output: list[dict] — 每个输入块展开为 [body_title_block, table_block]，
                         非表格块原样透传
    """
    from llm_router import route_table

    result = []
    for block in blocks:
        btype = block.get("type", "")
        if btype == "table":
            raw = json.dumps(block.get("content", []), ensure_ascii=False)
            result.extend(convert(route_table(raw)))
        elif btype == "body":
            result.extend(convert(route_table(block.get("content", ""))))
        else:
            result.append(block)
    return result
