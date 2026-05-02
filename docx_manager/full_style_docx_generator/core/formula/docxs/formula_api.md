# 公式检测与转换接口文档

## 数据流总览

### 方案 A：使用 hiagent 平台（原方案）

```
backfilled_styles.json
        ↓
[服务端] GET /detect-formulas
        规则检测 → 疑似公式列表
        ↓
[客户端 hiagent Handler 1] 获取疑似公式
        ↓
[客户端 LLM] 识别并提取公式内容
        ↓
[客户端 hiagent Handler 2] 发送确认公式列表
        ↓
[服务端] POST /convert-formulas
        latex → omath XML
        ↓
返回 omath 结果列表
```

### 方案 B：使用 BaseAgent（当前方案）

LLM 调用移至服务端，客户端只需触发一个路由，服务端完成全部三步。

```
backfilled_styles.json
        ↓
[服务端] GET /process-formulas
  ├─ Step 1: 规则检测 → 疑似公式列表
  │          detector.detect_formula_blocks()
  │
  ├─ Step 2: LLM 提取 → 确认公式列表
  │          base_agent.call_structured(system, user, FormulaListResponse)
  │          返回: [{id, text_before, latex_formula, text_after, label}, ...]
  │
  └─ Step 3: 转换 → omath
             converter.convert_formula_list()
        ↓
返回: [{id, text_before, omath, text_after, label}, ...]
```

**与方案 A 的对比**：

| | 方案 A (hiagent) | 方案 B (BaseAgent) |
|--|--|--|
| 客户端路由数 | 2 个 Handler | 0（仅触发） |
| LLM 调用位置 | 客户端 | 服务端 |
| 对外暴露接口 | `/detect-formulas` + `/convert-formulas` | `/process-formulas` |
| 灵活性 | 受平台限制 | 可控模型/参数/重试 |

---

## 阶段一：原始数据源

**文件**：`docx_manager/data/backfilled_styles.json`

```json
[
  {
    "id": "elem_1",
    "type": "heading1 | heading2 | heading3 | body",
    "content": "原始文本内容（纯文本，非固定格式）",
    "style": {
      "style_name": "string",
      "is_heading": false,
      "heading_level": 0
    }
  }
]
```

---

## 阶段二：服务端接口

### GET /detect-formulas

**描述**：读取 `backfilled_styles.json`，用规则检测疑似含公式的元素，返回候选列表供客户端 LLM 处理。

**请求**：无请求体

**响应 200**：
```json

  "suspected_formulas": [
    {
      "id": "elem_5",
      "content": "设损失函数为 L = \\frac{1}{n}\\sum_{i=1}^{n}(y_i - \\hat{y}_i)^2"
    }
  ]
```

**响应 404**：`backfilled_styles.json` 不存在
```json
{ "error": "No backfilled_styles.json found. Please run backfill-styles first." }
```

**检测规则（宽松，宁多勿漏）**：

| 规则 | 示例 |
|------|------|
| LaTeX 命令 `\frac \sum \int \sqrt` 等 | `\frac{a}{b}` |
| LaTeX 定界符 `$...$` 或 `$$...$$` | `$E=mc^2$` |
| 数学 Unicode 符号 `∑ ∫ ± × ÷ √ ∞` 等 | `∑xᵢ` |
| 上下标特征 `^` `_` 紧跟字母/数字 | `x^2`, `a_1` |
| 变量等式模式 `字母 = 数学表达式` | `E=mc^2` |

---

### POST /convert-formulas

**描述**：接收客户端 LLM 确认后的公式列表，将 `latex_formula` 转换为可直接插入 WPS/Word 的 OOXML `<m:oMath>`，原样透传其余字段。

**请求体**：
```json
{
  "formulas": [
    {
      "id": "elem_5",
      "text_before": "设损失函数为",
      "latex_formula": "L = \\frac{1}{n}\\sum_{i=1}^{n}(y_i - \\hat{y}_i)^2",
      "text_after": "，其中 n 为样本数。",
      "label": ""
    },
    {
      "id": "elem_9",
      "text_before": "",
      "latex_formula": "F = ma",
      "text_after": "",
      "label": "(4-1)"
    }
  ]
}
```

**字段说明**：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `id` | string | 是 | 对应 `backfilled_styles.json` 中的元素 id |
| `text_before` | string | 否 | 行内公式前的文本，块级公式时为空字符串 |
| `latex_formula` | string | 是 | LaTeX 公式内容，公式本身始终在此字段 |
| `text_after` | string | 否 | 行内公式后的文本，块级公式时为空字符串 |
| `label` | string | 否 | 块级公式的编号标记，如 `(4-1)`；行内公式时为空字符串 |

**行内公式 vs 块级公式区分**：

| 类型 | `text_before` / `text_after` | `label` |
|------|-------------------------------|---------|
| 行内公式 | 非空（公式嵌入句子中） | 空字符串 |
| 块级公式 | 空字符串 | 可能为 `(4-1)` 等，也可能为空 |

> 不会出现 `text_before`、`text_after`、`label` 三者同时非空的情况。

**响应 200**：
```json

  "results": [
    {
      "id": "elem_5",
      "text_before": "设损失函数为",
      "omath": "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">...</m:oMath>",
      "text_after": "，其中 n 为样本数。",
      "label": ""
    },
    {
      "id": "elem_9",
      "text_before": "",
      "omath": "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">...</m:oMath>",
      "text_after": "",
      "label": "(4-1)",
      "error": "[待确认] 仅在转换失败时出现，值为异常信息字符串"
    }
  ],
```

**响应字段说明**：

| 字段 | 类型 | 说明 |
|------|------|------|
| `omath` | string | 标准 OOXML `<m:oMath>` XML 字符串，可直接插入 docx |
| `error` | string | 仅转换失败时出现，`omath` 此时为空字符串 |
| `total` | int | 本次处理的公式总数 |
| `failed` | int | 转换失败的数量 |

**响应 400**：请求体缺少 `formulas`
```json
{ "error": "No formulas provided" }
```

---

## 阶段三：客户端 Handler 接口

### Handler 1 — 获取疑似公式

**输入 `params`**：

| 键 | 类型 | 说明 |
|----|------|------|
| `server_url` | string | 服务端地址，如 `http://localhost:5000` |

**输出**：

```json

  "suspected_formulas": [{ "id": "string", "content": "string" }],

```

---

### Handler 2 — 发送确认公式，获取 omath

**输入 `params`**：

| 键 | 类型 | 说明 |
|----|------|------|
| `server_url` | string | 服务端地址 |
| `formulas` | array | LLM 确认后的公式列表，结构见 POST /convert-formulas 请求体 |

**输出**：

```json

  "results": [
    {
      "id": "string",
      "text_before": "string",
      "omath": "string",
      "text_after": "string",
      "label": "string"
    }
  ]

```

---

## 转换技术链路

```
latex_formula (string)
    ↓ latex2mathml.converter.convert()
MathML (W3C 标准 XML)
    ↓ lxml XSLT (MML2OMML.XSL，微软官方)
<m:oMath> (OOXML，可插入 WPS/Word docx)
```

**XSL 文件位置**：`core/formula/assets/MML2OMML.XSL`（来源：Microsoft Office 安装目录）

**XSLT 对象全局缓存**：`converter.py` 中 `_transform` 为模块级单例，首次调用时初始化，避免重复解析 XSL。


