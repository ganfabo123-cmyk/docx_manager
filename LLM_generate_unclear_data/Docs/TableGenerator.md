# TableGenerator 文档

## 1. 模块概述

**文件路径：** `LLM_generate_unclear_data/table_generate.py`  
**职责：** 接收已确认的表格块列表，经大模型结构化后，输出对齐 `docx_parser` 标准格式的 JSON 块列表（每个表格展开为一个 body 标题块 + 一个 table 数据块）。

---

## 2. 数据结构

### `TableData` — 内部中间结构

```python
@dataclass
class TableData:
    title:   str               # 表格标题文字
    content: list[list[str]]   # 二维数组，第一行为列标题，其余为数据行
```

> `TableData` 仅作为内部中转，不对外暴露。最终输出为标准 JSON 块列表。

---

## 3. 输出格式

每个表格输入块最终展开为**两个**标准 JSON 块，格式与 `docx_parser` 完全一致：

**body 块（表格标题）：**
```json
{
  "type":    "body",
  "content": "表1 实验数据汇总",
  "style":   { "style_name": "Normal" }
}
```

**table 块（表格数据）：**
```json
{
  "type": "table",
  "content": [
    ["样本编号", "温度(℃)", "压力(MPa)"],
    ["S01", "25.3", "1.02"],
    ["S02", "26.1", "1.05"]
  ],
  "style": { "style_name": "Table", "rows": 3, "cols": 3 }
}
```

> - `content` 第一行为列标题/属性行，其余为数据行  
> - 无独立 `title` 字段，标题通过 body 块承载  
> - `style.rows` / `style.cols` 由程序自动计算

---

## 4. 函数接口

### `convert(table_json_str) → list[dict]`

**核心转换函数。** 将 LLM 返回的 JSON 解析并生成两个标准块。

```python
def convert(table_json_str: str) -> list[dict]:
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `table_json_str` | `str` | LLM 返回的 JSON 字符串 |

**期望的 LLM 输出 JSON 格式：**

```json
{
  "title": "表1 实验数据汇总",
  "content": [
    ["样本编号", "温度(℃)", "压力(MPa)"],
    ["S01", "25.3", "1.02"],
    ["S02", "26.1", "1.05"]
  ]
}
```

**返回值：** `list[dict]` — `[body_title_block, table_block]`

**异常：**
- `json.JSONDecodeError` — JSON 格式非法
- `KeyError` — 缺少必要字段 `content`

> **注意：** 此函数不调用大模型，是纯本地解析，可独立测试。

---

### `generate(blocks) → list[dict]`

**完整生成函数。** 接收已确认的表格块列表，逐块调用 LLM 转换，返回修改后的 JSON 块列表。

```python
def generate(blocks: list[dict]) -> list[dict]:
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `blocks` | `list[dict]` | 已确认的表格块列表，元素为 `type=="table"` 或 `type=="body"` |

**处理规则：**

| 输入块类型 | 处理方式 |
|-----------|---------|
| `type == "table"` | 将 `content`（二维数组）序列化为 JSON 字符串传给 LLM |
| `type == "body"` | 将 `content`（文本）直接传给 LLM |
| 其他 | 原样透传，不处理 |

**返回值：** `list[dict]` — 每个表格块展开为 `[body_title_block, table_block]`

**调用流程：**
```
blocks
    → 对每个 table/body 块：route_table(raw) → JSON 字符串 → convert()
    → 展开为 [body_title_block, table_block]
    → 合并后返回完整 list
```

---

## 5. 设计说明

- `content` 第一行固定为列标题行，与 `docx_parser` 的 table 块格式完全一致
- `title` 通过独立的 body 块承载，不存入 table 块，符合 Word 文档的语义结构
- `style.rows` / `style.cols` 由 `_make_blocks()` 自动从 `content` 计算，无需 LLM 提供
- 大模型生成的 JSON 结构严格符合 `docx_parser` 定义的标准，保证后续插入 Word 文档时的一致性
