# Detector 文档

## 1. 模块概述

**文件路径：** `LLM_generate_unclear_data/detector.py`  
**职责：** 接收 `docx_parser` 输出的 JSON 块列表，通过宽松规则进行粗筛，将块分类为疑似公式、疑似表格、疑似图片，再交由 LLM 细筛确认，最终输出已确认的分类结果供各 Generator 使用。

---

## 2. 输入格式

输入为 `docx_parser.py` 解析后的 JSON 块列表，每个块的结构如下：

**普通段落（body）：**
```json
{
  "id": "elem_1",
  "type": "body",
  "content": "当 x=0 时，F=ma，其中 m 是质量",
  "style": { "style_name": "Normal", "alignment": "justify", ... }
}
```

**公式块（formula）：**
```json
{
  "type": "formula",
  "label": "(4-2)",
  "omml": "<m:oMathPara>...</m:oMathPara>",
  "text_before": "由上述推导可得",
  "text_after": "其中 c 为光速",
  "is_inline": true
}
```

**OLE 公式块（formula，需转换）：**
```json
{
  "type": "formula",
  "label": "(4-2)",
  "omml": "",
  "ole_base64": "base64string",
  "image_base64": "base64string",
  "width_pt": 50.0,
  "height_pt": 20.0
}
```

**图片块（image）：**
```json
{
  "id": "elem_5",
  "type": "image",
  "base64": "base64string",
  "caption": "图 2-1  实验装置示意图",
  "width": 3.5,
  "height": 3.5,
  "position": "center"
}
```

**表格块（table）：**
```json
{
  "id": "elem_8",
  "type": "table",
  "content": [["列1", "列2"], ["val1", "val2"]],
  "style": { "style_name": "Table", "rows": 3, "cols": 2 }
}
```

---

## 3. 分类规则

### 3.1 固定分类（不经规则判断，直接归入）

| 块类型 | 归入分类 | 说明 |
|--------|----------|------|
| `type: "formula"`, `omml == ""` | **100%公式** | OLE 公式，无需交由LLM识别 |
| `type: "table"` | **100%表格** | 结构已确认，无需交由LLM识别|
| `type: "image"` | **疑似图片** | 需 LLM 分组并确定 anchor |
| `type: "formula"`, `omml != ""` | **100%公式** | 已是 OMath，无需处理 |
| `type: "heading*"` | **跳过** | 不在处理范围 |

### 3.2 宽松规则（针对 `type: "body"` 块）

**疑似公式** — 满足以下任意一条即标记：

| 条件 | 示例 |
|------|------|
| 含 `\`（LaTeX 命令特征） | `\frac`, `\sum`, `\alpha` |
| 含 `^` 或 `_` | `x^2`, `a_n` |
| 含数学符号 | `∑ ∫ ∏ √ ± × ÷ ≤ ≥ ≠ → ∈ ∀ ∃` |
| 含希腊字母 | `α β γ δ ε θ λ μ π σ φ ω` |
| 结尾匹配 `(数字-数字)` 模式 | ` F=ma (4-2)` |

**疑似表格** — 满足以下任意一条即标记：

| 条件 | 示例 |
|------|------|
| 含 `\t` 或连续 3 个以上空格 | 列对齐文本 |
| 含 `\|` | Markdown 表格风格 |
| 含换行符且各行字段数一致 | 多行结构化数据 |

---

## 4. 函数接口

### `detect(blocks) → dict`

**粗筛函数。** 遍历所有块，应用分类规则，返回疑似项分组。

```python
def detect(blocks: list[dict]) -> dict:
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `blocks` | `list[dict]` | `docx_parser` 输出的 JSON 块列表 |

**返回值：**
```python
{
    "formula": [...]，  # 疑似公式块列表
    "table":   [...],  # 疑似表格块列表
    "image":   [...]   # 疑似图片块列表
}
```

---

### `confirm(detected) → dict`

**细筛函数。** 将粗筛结果送 LLM 确认，过滤误判，返回已确认的分类结果。

```python
def confirm(detected: dict) -> dict:
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `detected` | `dict` | `detect()` 的返回值 |

**返回值：** 与 `detect()` 结构相同，但经 LLM 过滤后仅保留真正需要处理的块

**调用流程：**
```
detected
    → route_confirm(detected)   # llm_router.py 中的路由函数
    → LLM 判断每个疑似项是否为真正的公式 / 表格 / 图片
    → 返回过滤后的 dict
```

---

## 5. llm_router.py 新增函数

### `route_confirm(detected) → dict`

```python
def route_confirm(detected: dict) -> dict:
```

**大模型任务：**
- 对疑似公式：判断 body 块的文本内容是否真正包含数学公式
- 对疑似表格：判断 body 块的文本内容是否真正描述了一个表格
- 对疑似图片：图片块无需二次确认，直接透传（LLM 主要处理分组和 anchor）

**返回值：** 过滤后的 `dict`，结构同 `detect()` 输出

---

## 6. 整体调用流程

```
docx_parser 输出的 JSON 块列表
        ↓
  detect(blocks)          ← 宽松规则粗筛
        ↓
  confirm(detected)       ← LLM 细筛（调用 route_confirm）
        ↓
  confirmed["formula"]  → FormulaGenerator.generate()
  confirmed["table"]    → TableGenerator.generate()
  confirmed["image"]    → ImagePositionGenerator.generate()
```
