# hit-paper-helper 开发笔记

## 今日 OKR（2026-04-18）

**O：实现论文图片自动排版**

| KR | 目标 | 状态 |
|----|------|------|
| KR1 | 单图片完整排版流程跑通 | 🔄 进行中 |
| KR2（plus）| 双图/三图并列排版 | ⬜ 待开始 |

---

## 当前进度

### 已完成

- `primitives.py` — 原子操作层，新增 `locate_image_box` / `find_and_click_image`（支持 region 限定搜索范围，PIL 读图绕过 OpenCV 中文路径问题）
- `wps_nav.py` — 导航层
  - OCR reader 加入 `en` 语言，解决英文 marker 识别失败问题
  - `navigate_to_crop_inputs()`：模板匹配依次定位属性面板 → 图片 → 裁剪 → 宽度/高度输入框；第一步全屏定位面板区域，后续步骤限定在该区域内防误匹配
- `insert_image.py` — 工作流层
  - `_locate_caption_line()`：OCR 找 MARKER 后鼠标上移偏置并点击，定位图题行
  - `open_image_property_panel()`：Shift+F10 → O 打开属性面板
  - `insert_image_after_paragraph()`：主流程串联上述步骤

### 关键设计决策

- **MARKER 识别尾部多 I**：光标闪烁竖线被 OCR 识别为 `I`，用 `in` 子串匹配规避
- **模板匹配防误匹配**：第一个模板全屏找，锁定面板列区域，后续全部在该 region 内搜索
- assets 图片命名：`panel_tab_image` / `panel_image_btn` / `panel_image_crop` / `panel_crop_width` / `panel_crop_height`

### 待完成

- 宽度/高度数值输入 + 确认
- 删除 MARKER 文本
- 保存关闭
- 双图/三图并列排版逻辑
