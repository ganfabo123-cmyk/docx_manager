# LLMRouter 文档

## 1. 模块概述

**文件路径：** `LLM_generate_unclear_data/llm_router.py`  
**职责：** 作为业务模块（FormulaGenerator、TableGenerator、ImagePositionGenerator）与大模型通信层（BaseAgent）之间的路由中间层。每个路由函数负责为对应业务构造 Prompt，调用 BaseAgent，并返回原始响应文本。

---

## 2. 设计定位

```
FormulaGenerator  ──┐
TableGenerator    ──┼──→  llm_router.py  ──→  BaseAgent  ──→  远程大模型平台
ImagePositionGenerator ──┘
```

**路由层的职责边界：**

| 职责 | 归属 |
|------|------|
| Prompt 内容构造 | `llm_router.py` |
| HTTP 通信 | `BaseAgent` |
| 响应文本解析（JSON → 数据结构） | 各 Generator 的 `convert()` |
| 业务逻辑 | 各 Generator 的 `generate()` |

---

## 3. 函数接口

### `route_formula(raw_text) → list[FormulaItem]`

为公式结构化任务构造 Prompt，调用大模型，返回解析后的公式列表。

```python
def route_formula(raw_text: str) -> list[FormulaItem]:
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `raw_text` | `str` | 原始混合段落字符串（普通文本 + 混乱公式） |

**返回值：** `list[FormulaItem]`（`InlineFormula` 或 `BlockFormula` 的混合列表）

**大模型任务：**
- 识别段落中的所有公式
- 判断每个公式是行内公式（InlineFormula）还是独立公式（BlockFormula）
- 提取 `text_before`、`formula`、`text_after` 或 `prefix`、`formula`
- 将 `formula` 规范化为标准 LaTeX

**期望大模型输出格式（JSON）：**
```json
[
  {
    "type": "inline",
    "text_before": "由上述推导可得 ",
    "formula": "E=mc^2",
    "text_after": "，其中 c 为光速"
  },
  {
    "type": "block",
    "prefix": "4-2",
    "formula": "F = ma"
  }
]
```

---

### `route_table(raw_input) → str`

为表格结构化任务构造 Prompt，调用大模型，返回 JSON 字符串（由 `TableGenerator.convert()` 解析）。

```python
def route_table(raw_input: str) -> str:
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `raw_input` | `str` | 原始混乱表格文本，格式不限 |

**返回值：** `str` — JSON 字符串，格式符合 `TableData` 结构

**大模型任务：**
- 识别表格标题
- 提取各列属性名（表头）
- 提取所有数据行
- 统一单位格式，归入属性名中

**期望大模型输出格式（JSON）：**
```json
{
  "title": "表1 实验数据汇总",
  "attributes": ["样本编号", "温度(℃)", "压力(MPa)"],
  "rows": [
    ["S01", "25.3", "1.02"],
    ["S02", "26.1", "1.05"]
  ]
}
```

---

### `route_image_position(images, paragraphs) → str`

为图片定位任务构造 Prompt，调用大模型，返回 JSON 字符串（由 `ImagePositionGenerator.convert()` 解析）。

```python
def route_image_position(images: list, paragraphs: list[str]) -> str:
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `images` | `list` | 图片列表，元素格式**[待确认]** |
| `paragraphs` | `list[str]` | 文档各段落文本，顺序与文档一致 |

**返回值：** `str` — JSON 字符串，格式符合 `ImageGroup` 列表结构

**大模型任务：**
- 理解各图片内容（多模态）
- 分析图片与段落的语义关联
- 对图片进行合理分组
- 为每组图片找到最合适的锚点段落，返回关键文字片段

**期望大模型输出格式（JSON）：**
```json
[
  {
    "image_indices": [0, 1],
    "anchor_text": "图1和图2展示了装置的正视图与侧视图"
  },
  {
    "image_indices": [2, 3],
    "anchor_text": "如图3、图4所示，温度分布呈现梯度变化"
  }
]
```

---

## 4. 设计说明

- 路由函数内部不做业务判断，只负责 Prompt 构造和 BaseAgent 调用
- `route_formula()` 直接返回解析后的 `list[FormulaItem]`（因为需要根据 `type` 字段分支构造不同数据类）；其余两个路由函数返回原始 JSON 字符串，由各自的 `convert()` 负责解析
- Prompt 的具体内容（system prompt 措辞、few-shot 示例等）在实现阶段确定，文档不作约定
