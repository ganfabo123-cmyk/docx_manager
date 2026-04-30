# WPS UI 自动化函数速查

> 分三层：primitives（原子）→ wps_nav（WPS操作）→ workflows/insert_image（排版工作流）

---

## Layer 1 · primitives.py — 原子操作

| 函数 | 作用 |
|------|------|
| `wait(t=0.25)` | 等待 t 秒 |
| `press(key)` | 按单键 |
| `hotkey(*keys)` | 按组合键，如 `hotkey('ctrl','s')` |
| `type_text(text)` | 剪贴板粘贴文字（支持中文） |
| `click(x, y)` | 左键点击坐标 |
| `right_click(x, y)` | 右键点击坐标 |
| `move(x, y)` | 移动鼠标（不点击） |
| `screenshot()` | 全屏截图，返回 PIL Image |
| `locate_image_box(template_path, confidence, region)` | 模板匹配，返回 `(left, top, w, h)`，找不到抛异常 |
| `find_and_click_image(template_path, confidence, move_only, region)` | 模板匹配并点击中心，返回 `(cx, cy)` |

---

## Layer 2 · wps_nav.py — WPS 操作

### 文档生命周期

| 函数 | 快捷键 / 操作 | 作用 |
|------|--------------|------|
| `open_doc(path, wait_sec=4)` | `os.startfile` | 用 WPS 打开文档 |
| `save()` | `Ctrl+S` | 保存 |
| `close()` | `Alt+F4` | 关闭窗口 |
| `save_close()` | — | 保存后关闭 |

### 光标导航

| 函数 | 快捷键 | 作用 |
|------|--------|------|
| `goto_start()` | `Ctrl+Home` | 跳文档开头 |
| `goto_end()` | `Ctrl+End` | 跳文档末尾 |
| `goto_line_end()` | `End` | 跳行尾 |
| `newline()` | `Enter` | 插入换行 |
| `jump_next_section()` | `Ctrl+G → Alt+T → Esc` | 跳到下一节 |

### 文本查找

| 函数 | 快捷键 | 作用 |
|------|--------|------|
| `find_text(text)` | `Ctrl+F → 输入 → Enter → Esc` | 查找文本，光标停在匹配处 |

### 页码对话框

| 函数 | 作用 |
|------|------|
| `open_page_number_dialog()` | 打开页码格式对话框（Alt+P+N+U+N） |
| `page_dialog_move(down, up)` | 对话框内方向键移动 |
| `page_dialog_apply_this_section_onward()` | 应用范围=本节及以后（Alt+P） |
| `confirm()` | 确认（Enter） |

### 插入图片对话框

| 函数 | 作用 |
|------|------|
| `open_insert_picture_dialog()` | Alt+N+P+P 打开插入图片对话框 |
| `input_file_path_confirm(path)` | 地址栏粘贴路径并两次 Enter 确认 |

### 图片属性面板（模板匹配）

| 函数 | 作用 |
|------|------|
| `navigate_to_crop_inputs(img_width, img_height, click_crop=True)` | 在已打开属性面板中依次点击图片→裁剪→填写宽高，并将裁剪位置宽高同步、偏移XY归零 |

> `click_crop=False`：跳过点击"裁剪"步骤（第2张图起属性面板已在裁剪页时用）

### 图片感知（屏幕差分 / OCR）

| 函数 | 作用 |
|------|------|
| `capture_screen()` | 截图（插图前调用，供差分用） |
| `get_changed_center(before)` | 差分定位图片出现的中心坐标 `(cx, cy)` |
| `find_and_click_text(text, region)` | OCR 识别屏幕文字并点击，`region=(left,top,w,h)` 可缩小范围 |
| `set_picture_size_via_panel(cx, cy, width_cm, height_cm)` | OCR方式：右键→属性→图片→裁剪→填宽高（较慢，备用） |

---

## Layer 3 · workflows/insert_image.py — 图片排版工作流

### 可复用辅助函数

| 函数 | 作用 |
|------|------|
| `_find_image_center(image_path, confidence=0.09)` | pyautogui 模板匹配定位图片中心，返回 `(cx, cy)` |
| `open_image_property_panel(x, y)` | 右键点击图片 → 按 O，打开属性面板，等待3秒 |
| `_set_align_center_ribbon()` | ribbon 操作：`alt+H+A+L → alt+H+A+C`（WPS居中对齐序列） |
| `_format_caption_line()` | 选中当前行 → 居中对齐 → 段前距清零（图题专用） |
| `_clean_image_line(img_cx, img_cy)` | 点击图片 → 左移 → 删2个幽灵字符 → 居中对齐 |

### 主工作流

| 函数 | 作用 |
|------|------|
| `insert_image_after_paragraph(docx_path, anchor_text, image_path, caption, close_after=True, cnt=0)` | 在 anchor_text 段落后插入单图+图题+设置尺寸；`close_after=False` 时只保存不关闭；`cnt>0` 时跳过属性面板裁剪tab点击 |
| `insert_n_images_one_col(docx_path, anchor_text, items)` | 批量单列插入，`items=[(image_path, caption), ...]`；第2轮起自动以上一轮 caption 为 anchor |

---

## 常量

| 位置 | 常量 | 默认值 | 说明 |
|------|------|--------|------|
| `insert_image.py` | `HIT_DEFAULT_SINGLE_IMG_WIDTH` | 12.00 cm | 单图宽度 |
| `insert_image.py` | `HIT_DEFAULT_SINGLE_IMG_HEIGHT` | 6.00 cm | 单图高度 |
| `insert_two_images.py` | `HIT_DEFAULT_DOUBLE_IMG_WIDTH` | 6.99 cm | 双列图宽度 |
| `insert_two_images.py` | `HIT_DEFAULT_DOUBLE_IMG_HEIGHT` | 4.99 cm | 双列图高度 |
| `calibration.py` | `TEXT_LEFT_X` | 待填 | 文本区左边界屏幕X |
| `calibration.py` | `SPACE_WIDTH_PX` | 待填 | 半角空格像素宽 |
| `calibration.py` | `CHAR_WIDTH_PX` | 待填 | 标题字符平均像素宽 |

---

## 典型调用链（开发新排版时参考）

```
open_doc → goto_start → find_text → goto_line_end
→ newline → open_insert_picture_dialog → input_file_path_confirm
→ _find_image_center → open_image_property_panel → navigate_to_crop_inputs
→ _clean_image_line / _format_caption_line
→ save_close
```
