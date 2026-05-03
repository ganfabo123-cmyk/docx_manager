import re

_PIPE_TABLE = re.compile(r'\|.+\|')
_SEPARATOR_ROW = re.compile(r'^\s*\|?[\s\-:=+|]{3,}\|?\s*$')
_TABLE_TITLE_RE = re.compile(r'^(表|Table)\s*\d', re.IGNORECASE)


def is_suspected_table(content: str) -> bool:
    if not content or not isinstance(content, str) or not content.strip():
        return False

    # 竖线表格（Markdown pipe table，允许不规则）
    if _PIPE_TABLE.search(content):
        return True

    # 制表符分列
    if '\t' in content:
        return True

    # 多行 + 多列结构：去掉纯分隔行后，各行用2+空格拆分，列数基本一致
    lines = [
        l.strip() for l in content.split('\n')
        if l.strip() and not _SEPARATOR_ROW.match(l)
    ]
    if len(lines) >= 2:
        counts = [len(re.split(r'\s{2,}', l)) for l in lines]
        valid = [c for c in counts if c >= 2]
        if len(valid) >= 2 and len(set(valid)) <= 2:
            return True

    return False


def is_table_title(content: str) -> bool:
    if not content or not isinstance(content, str):
        return False
    return bool(_TABLE_TITLE_RE.match(content.strip()))


def detect_table_blocks(elements: list) -> list:
    return [
        {"id": elem["id"], "content": elem["content"]}
        for elem in elements
        if is_suspected_table(elem.get("content", ""))
    ]


def group_table_blocks(elements: list, suspected: list) -> list[list[dict]]:
    """按在原始 elements 列表中的连续位置分组，同一表格的行归为一组。"""
    if not suspected:
        return []

    id_to_pos = {elem["id"]: i for i, elem in enumerate(elements)}
    suspected_sorted = sorted(suspected, key=lambda x: id_to_pos.get(x["id"], float('inf')))

    groups = [[suspected_sorted[0]]]
    for i in range(1, len(suspected_sorted)):
        prev_pos = id_to_pos.get(suspected_sorted[i - 1]["id"], -1)
        curr_pos = id_to_pos.get(suspected_sorted[i]["id"], -1)
        if curr_pos == prev_pos + 1:
            groups[-1].append(suspected_sorted[i])
        else:
            groups.append([suspected_sorted[i]])

    return groups
