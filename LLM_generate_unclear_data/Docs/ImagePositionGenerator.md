
---

# ImagePositionGenerator 文档

## 1. 模块概述

**文件路径：** `LLM_generate_unclear_data/image_position_generate.py`  
**职责：** 接收图片列表与文档段落列表（由文档解析所得），经大模型分析后，输出图片的分组方案、组内排序方案及每组图片在文档中的锚点段落索引（`anchor_idx`）。

---

## 2. 数据结构

### `ImageGroup` — 图片分组与锚点

```python
@dataclass
class ImageGroup:
    image_indices: list[int]   # 本组图片在传入列表中的索引（从 0 开始）
    anchor_idx:    int         # 锚点段落在文档段落列表中的索引（即插入位置）
```

**示例：**

传入图片：`[图1.png, 图2.png, 图3.png, 图4.png]`  
传入段落：`["段落0内容...", "段落1：图1和图2展示了...", "段落2：...", "段落3：如图3、图4所示..."]`

结构化结果：
```python
[
    ImageGroup(
        image_indices = [0, 1],
        anchor_idx    = 1
    ),
    ImageGroup(
        image_indices = [2, 3],
        anchor_idx    = 3
    )
]
```

> **说明：** `anchor_idx` 指示了图片组应插入到文档段落列表中的位置（在该索引段落之后）。

---

## 3. 函数接口

### `group_images(images, paragraphs) → list[list[int]]`

**图片分组函数。** 调用大模型分析图片内容与段落语义，确定哪些图片属于同一组。

```python
def group_images(images: list, paragraphs: list[str]) -> list[list[int]]:
```

### `sort_group_images(image_indices, paragraphs) → list[int]`

**组内排序函数。** 根据文档语义，确定组内图片的逻辑排列顺序。

```python
def sort_group_images(image_indices: list[int], paragraphs: list[str]) -> list[int]:
```

### `convert(position_json_str) → list[ImageGroup]`

**解析函数。** 将大模型返回的 JSON 字符串解析为 `ImageGroup` 列表。

```python
def convert(position_json_str: str) -> list[ImageGroup]:
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `position_json_str` | `str` | JSON 字符串，由大模型生成 |

**期望的 JSON 格式：**

```json
[
  {
    "image_indices": [0, 1],
    "anchor_idx": 1
  },
  {
    "image_indices": [2, 3],
    "anchor_idx": 3
  }
]
```

### `generate(images, paragraphs) → list[ImageGroup]`

**完整生成函数。** 接收图片列表与段落列表，协调上述函数完成最终结构化输出。

```python
def generate(images: list, paragraphs: list[str]) -> list[ImageGroup]:
```

**调用流程：**
1. `group_images(...)`：调用大模型获取初步分组。
2. 对每组调用 `sort_group_images(...)`：调用大模型确定组内图片的顺序。
3. 调用大模型获取每组对应的 `anchor_idx`。
4. 封装为 `ImageGroup` 对象列表并返回。

---

## 4. 设计说明

- 分组与排序策略由大模型根据图片内容与段落语义自动判断。
- `image_indices` 中的索引对应 `generate()` 传入的 `images` 列表初始顺序。
- `anchor_idx` 指向段落列表的下标，确保程序能准确将图片对象插入到列表的对应位置。
- 该程序假设接收之前文档解析生成的 `json` 列表作为输入基础。