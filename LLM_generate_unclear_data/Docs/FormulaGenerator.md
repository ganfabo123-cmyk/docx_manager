# FormulaGenerator 文档

## 1. 模块概述

**文件路径：** `LLM_generate_unclear_data/formula_generate.py`  
**职责：** 接收含有混乱公式的原始段落文本，经大模型结构化后，将公式从 LaTeX 转换为可嵌入 Word 文档的 OMath XML 格式。

---

## 2. 数据结构

### `InlineFormula` — 行内公式

公式位于段落文本中，前后伴随普通文字。

```python
@dataclass
class InlineFormula:
    text_before: str   # 公式前的文字片段
    formula: str       # 公式内容，LaTeX 格式
    text_after: str    # 公式后的文字片段
```

**示例：**

原始段落：`"由上述推导可得 E=mc^2，其中 c 为光速"`

结构化结果：
```
InlineFormula(
    text_before = "由上述推导可得 ",
    formula     = "E=mc^2",
    text_after  = "，其中 c 为光速"
)
```

---

### `BlockFormula` — 独立公式（带编号）

公式独立成行，带有编号前缀（如 `"4-2"`）。此类公式不属于段落行内内容。

```python
@dataclass
class BlockFormula:
    prefix: str    # 公式编号，如 "4-2"、"(1)" 等
    formula: str   # 公式内容，LaTeX 格式
```

**示例：**

原始文本：`"F = ma   (4-2)"`

结构化结果：
```
BlockFormula(
    prefix  = "4-2",
    formula = "F = ma"
)
```

---

### `FormulaItem`

```python
FormulaItem = Union[InlineFormula, BlockFormula]
```

---

## 3. 函数接口

### `convert(formula_items) → [待确认]`

**核心转换函数。** 将已结构化公式列表中的 `formula` 字段从 LaTeX 转换为 OMath XML。

```python
def convert(formula_items: list[FormulaItem]):
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `formula_items` | `list[FormulaItem]` | 已结构化的公式列表，LaTeX 格式 |

**返回值：** `[json块中的标准formula类型文本块]` 

> **注意：** 此函数不调用大模型，是纯本地转换，可独立测试。

---

### `generate(raw_text) → [json块中的标准formula类型文本块]`

**完整生成函数。** 接收原始段落，通过路由层调用大模型完成结构化，再调用 `convert()`。

```python
def generate(raw_text: str):
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `raw_text` | `str` | 原始段落字符串，包含混乱的普通文本与公式符号 |

**返回值：** 同 `convert()`

**调用流程：**
```
raw_text
    → route_formula(raw_text)     # 调用大模型，返回 list[FormulaItem]
    → convert(formula_items)      # 本地转换 LaTeX → OMath
    → 返回结果
```

---

## 4. 设计说明

- `convert()` 只负责 LaTeX → OMath 的机械转换，语义理解由大模型完成
- 两种公式类型（`InlineFormula` / `BlockFormula`）共享同一个 `formula` 字段，`convert()` 的核心逻辑对两者一致
- `text_before` / `text_after` 用于后续还原公式在段落中的位置，当前阶段暂不使用
- `prefix` 仅用于标记编号，不参与公式转换
