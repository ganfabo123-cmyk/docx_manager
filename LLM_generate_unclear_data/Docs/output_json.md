# 文档 JSON 块结构参考

> 来源：`docx_manager/utils/docx_parser.py`  
> 所有块均为 `DocxParser.parse()` 返回列表中的一项。  
> 标注 `*可选` 的字段仅在满足条件时出现。

---

## 1. body — 普通段落

```json
{
  "id": "elem_1",
  "type": "body",
  "content": "段落的完整文本内容",
  "style": {
    "style_name": "Normal",
    "alignment": "justify",
    "spacing_before": 0,
    "spacing_after": 0,
    "line_spacing": 480,
    "first_line_indent": 420,
    "left_indent": 0,
    "is_heading": false,
    "heading_level": 0
  }
}
```

> `alignment` 取值：`"left"` / `"center"` / `"right"` / `"justify"`  
> `spacing_*`、`line_spacing`、`first_line_indent`、`left_indent` 均为 *可选*，仅在文档中有对应设置时出现。

---

## 2. heading — 标题段落

```json
{
  "id": "elem_2",
  "type": "heading1",
  "content": "第一章 引言",
  "style": {
    "style_name": "Heading 1",
    "alignment": "left",
    "is_heading": true,
    "heading_level": 1
  }
}
```

> `type` 取值：`"heading1"` / `"heading2"` / `"heading3"` ...  
> `heading_level` 与 `type` 末尾数字一致。

---

## 3. formula — 公式

### 3a. 独立公式（块级，oMathPara）

```json
{
  "type": "formula",
  "label": "(4-2)",
  "omml": "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">...</m:oMathPara>",
  "text_before": "由上述推导可得",
  "text_after": "其中 c 为光速"
}
```

> - **无 `id` 字段**（与其他类型不同）  
> - `label`：公式编号如 `"(4-2)"`，无编号则为 `""`  
> - `text_before` / `text_after` *可选*，仅当公式前后有文字时出现  
> - 无 `is_inline` 字段

### 3b. 行内公式（oMath）

```json
{
  "type": "formula",
  "label": "",
  "omml": "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">...</m:oMath>",
  "is_inline": true,
  "text_before": "满足条件",
  "text_after": "时成立"
}
```

> - `is_inline: true` 仅行内公式有此字段  
> - `text_before` / `text_after` *可选*

### 3c. OLE 公式（旧版 Equation Editor，omml 为空）

```json
{
  "type": "formula",
  "label": "(4-2)",
  "omml": "",
  "ole_base64": "base64编码的OLE二进制数据",
  "prog_id": "Equation.DSMT4",
  "image_base64": "base64编码的公式预览图",
  "width_pt": 50.0,
  "height_pt": 20.0,
  "text_before": "由此得",
  "text_after": ""
}
```

> - `omml` 为空字符串是识别 OLE 公式的标志  
> - `image_base64` *可选*，存在时可用于多模态识别  
> - `width_pt` / `height_pt` *可选*

---

## 4. image — 嵌入图片

```json
{
  "id": "elem_5",
  "type": "image",
  "base64": "base64编码的图片数据",
  "caption": "图 2-1  实验装置整体示意图",
  "width": 3.5,
  "height": 3.5,
  "position": "center"
}
```

> - 只处理嵌入型（`wp:inline`），跳过浮动型锚点图  
> - `caption` 由全文图题预扫描填入，无图题则为 `""`  
> - `width` / `height` 单位为英寸，*可选*（无尺寸信息时可能为 `null`）  
> - `position` 当前固定为 `"center"`

---

## 5. table — 表格

```json
{
  "id": "elem_8",
  "type": "table",
  "content": [
    ["样本编号", "温度(℃)", "压力(MPa)"],
    ["S01", "25.3", "1.02"],
    ["S02", "26.1", "1.05"]
  ],
  "style": {
    "style_name": "Table",
    "rows": 3,
    "cols": 3,
    "width": 5000
  }
}
```

> - `content` 为二维字符串数组，**第一行为列标题/属性行**，其余为数据行  
> - **无独立 `title` 字段**，表格标题通常是表格前后的 body 段落  
> - `style.width` *可选*，单位 twips（1/1440 英寸）

---

## 汇总：各类型字段对照

| 字段 | body | heading | formula | image | table |
|------|------|---------|---------|-------|-------|
| `id` | ✓ | ✓ | **✗** | ✓ | ✓ |
| `type` | `"body"` | `"headingN"` | `"formula"` | `"image"` | `"table"` |
| `content` | 字符串 | 字符串 | — | — | 二维数组 |
| `style` | ✓ | ✓ | — | — | ✓ |
| `omml` | — | — | ✓ | — | — |
| `label` | — | — | ✓ | — | — |
| `is_inline` | — | — | 可选 | — | — |
| `text_before` | — | — | 可选 | — | — |
| `text_after` | — | — | 可选 | — | — |
| `ole_base64` | — | — | 可选 | — | — |
| `base64` | — | — | — | ✓ | — |
| `caption` | — | — | — | ✓ | — |
| `width` / `height` | — | — | — | 可选 | — |
