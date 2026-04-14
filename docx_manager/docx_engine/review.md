
# 缺陷修复与代码修改说明文档

## 一、 已解决问题：公式索引（Formula Index）显示异常
*   **问题原因：** 原程序逻辑从 `elem` 中获取的字段名为 `formula_index`，而实际 `extraction.json` 中存储该数据的字段名为 `formula`，导致数据对应失败。
*   **当前状态：** **已修复**。该问题已由提交人完成代码修改，目前公式索引已可正常显示，无需开发人员再作处理。

## 二、 待修改问题：OLE 预览图显示异常
*   **问题原因：** 目前程序逻辑尝试从 `elem` 中获取 `image_base64` 字段作为预览图的数据来源。但是，根据 `extraction.json` 的实际数据结构，节点中仅提供了一个 `base64` 字段，并不存在 `image_base64` 字段。
*   **修改要求：** 
    需要调整获取预览图的逻辑：**请基于 `extraction.json` 中现有的 `base64` 字段来生成预览图（image_base64）**，而非直接从 `elem` 中获取 `image_base64`。

## 三、 参考数据与代码

### 1. 当前 `docx_compiler` 中的图片处理代码（需修改部分）
```python
        # Optional WMF preview image (keeps static preview when OLE host inactive)
        img_b64 = elem.get('image_base64', '')
        if img_b64:
            try:
                img_bytes = base64.b64decode(img_b64, validate=True)
                img_rid   = self._embed_image(img_bytes, 'wmf', work_dir)
                imgdata   = ET.SubElement(shape, f'{{{V_NS}}}imagedata')
                imgdata.set(f'{{{R}}}id',       img_rid)
                imgdata.set(f'{{{O_NS}}}title', '')
            except Exception:
                pass  # preview is optional; skip silently on any error\
```

### 2. 实际的 `extraction.json` OLE 节点数据结构示例
```json
    {
      "index": 46,
      "type": "ole",
      "formula": "(4-1)",
      "formula_index": "",
      "position": "center",
      "suffix_position": "right",
      "base64": "0M.(此处忽略一堆字符)A",
      "width_pt": 81.9,
      "height_pt": 38.0,
      "prog_id": "Equation.3"
    }
```