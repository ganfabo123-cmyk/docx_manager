"""
Layer 3 — 双图横排排版工作流（ArUco 锚定图版）。

使用一张 ArUco 锚定图作为唯一 CV2 识别目标，
用户图片全程通过 Tab / Down 键盘跳转操作，
避免因图片内容/颜色/尺寸差异导致识别失败。

输入：
  docx_path     : 文档路径
  anchor_text   : 定位段落文字
  anchor_image  : ArUco 锚定图绝对路径（仅用于 CV2 定位，写死测试）
  images        : 用户图片路径列表
  captions      : 每张图的子图题，与 images 等长
  total_caption : 总图题
"""
from .. import wps_nav as W
from .. import primitives as P

HIT_DEFAULT_DOUBLE_IMG_WIDTH  = 6.99
HIT_DEFAULT_DOUBLE_IMG_HEIGHT = 4.99


def _locate_anchor_with_scroll(
    anchor_image: str,
    confidence: float = 0.7,
    max_scrolls: int = 15,
    scroll_amount: int = -3,
) -> tuple[int, int]:
    """屏幕上定位 ArUco 锚定图；找不到则向下滚动后重试。"""
    import pyautogui
    for _ in range(max_scrolls):
        try:
            cx, cy, _, _ = P._find_image_center_cv2(anchor_image, confidence=confidence)
            return cx, cy
        except RuntimeError:
            pyautogui.scroll(scroll_amount)
            P.wait(0.5)
    raise RuntimeError(f"滚动 {max_scrolls} 次后仍未找到锚定图：{anchor_image}")


def _open_property_panel() -> None:
    """当前图片已被 Tab 选中时，通过键盘打开属性面板：Shift+F10 → O"""
    import pyautogui
    pyautogui.hotkey('shift', 'f10')
    P.wait(0.2)
    P.press('o')
    P.wait(0.5)


def _set_center_align() -> None:
    """先左对齐再居中，确保任意初始格式都能正确居中。"""
    P.hotkey('alt')
    P.wait(0.1)
    P.hotkey('A')
    P.wait(0.1)
    P.hotkey('L')
    P.wait(0.1)
    P.hotkey('alt')
    P.wait(0.1)
    P.hotkey('A')
    P.wait(0.1)
    P.hotkey('C')
    P.wait(0.1)


def _insert_image(path: str) -> None:
    W.open_insert_picture_dialog()
    W.input_file_path_confirm(path)
    P.wait(0.3)
    P.hotkey('right')  # 图片插入后处于选中态，right 移至图片后文本位置


def insert_n_images_two_col(
    docx_path: str,
    anchor_text: str,
    anchor_image: str,
    images: list[str],
    captions: list[str],
    total_caption: str,
) -> None:
    n = len(images)
    num_rows = (n + 1) // 2
    labels = [f"({chr(ord('a') + i)})" for i in range(n)]

    # ── Phase 1: 插入所有内容 ──────────────────────────────────────────────
    W.open_doc(docx_path)
    W.goto_start()
    P.wait(4)

    print(f"→ 定位锚定文本：{anchor_text!r}")
    W.find_text(anchor_text)
    W.goto_line_end()

    print("→ 插入 ArUco 锚定图")
    P.hotkey('enter')
    _insert_image(anchor_image)  # 插入后 cursor 在锚定图后

    print(f"→ 插入 {n} 张用户图片（两列布局）")
    for i in range(0, n, 2):
        # 标签行
        P.hotkey('enter')
        P.type_text(labels[i])
        if i + 1 < n:
            P.type_text(labels[i + 1])
        # 图片行
        P.hotkey('enter')
        _insert_image(images[i])
        if i + 1 < n:
            _insert_image(images[i + 1])

    print("→ 插入子图题和总图题")
    P.hotkey('enter')
    for idx, cap in enumerate(captions):
        P.type_text(f"{labels[idx]} {cap}")
        P.hotkey('enter')
    P.type_text(total_caption)
    P.wait(0.3)

    # ── Phase 2: 调整图片尺寸 ─────────────────────────────────────────────
    print("→ Ctrl+F 定位锚定文本，准备调整图片尺寸")
    W.find_text(anchor_text)
    P.wait(0.5)

    print("→ CV2 定位锚定图（自动向下滚动）")
    anchor_cx, anchor_cy = _locate_anchor_with_scroll(anchor_image)
    P.click(anchor_cx, anchor_cy)
    P.wait(0.3)

    print(f"→ Tab 逐图调整尺寸，共 {n} 张")
    for i in range(n):
        print(f"   图 {i + 1}/{n}")
        P.hotkey('tab')
        P.wait(0.3)
        _open_property_panel()
        W.navigate_to_crop_inputs(
            img_width=HIT_DEFAULT_DOUBLE_IMG_WIDTH,
            img_height=HIT_DEFAULT_DOUBLE_IMG_HEIGHT,
            crop_width=HIT_DEFAULT_DOUBLE_IMG_WIDTH,
            crop_height=HIT_DEFAULT_DOUBLE_IMG_HEIGHT,
            click_crop=(i == 0),
        )
        P.wait(0.5)

    # ── Phase 3: 调整居中对齐 ─────────────────────────────────────────────
    print("→ 重新定位锚定图，准备调整居中")
    W.find_text(anchor_text)
    P.wait(0.5)
    anchor_cx, anchor_cy = _locate_anchor_with_scroll(anchor_image)
    P.click(anchor_cx, anchor_cy)
    P.wait(0.3)

    # Tab 一次跳到第一张用户图（奇数行首图）
    P.hotkey('tab')
    P.wait(0.3)

    print(f"→ 对 {num_rows} 行图片调整居中")
    for row in range(num_rows):
        print(f"   行 {row + 1}/{num_rows}")
        # 在两图之间加 4 个空格撑开间距，right 先移出图片选中态
        P.hotkey('right')
        P.type_text('    ')
        # 选中整行（含两张图）
        P.hotkey('home')
        P.wait(0.1)
        P.hotkey('shift', 'end')
        P.wait(0.1)
        _set_center_align()
        # Tab 在对齐操作后失效，改用 Down 移到下一行图片
        # 需按两次：第一次跳过标签行，第二次到图片行
        if row < num_rows - 1:
            P.hotkey('down')
            P.wait(0.1)
            P.hotkey('down')
            P.wait(0.2)

    # ── Phase 4: 图题居中 ─────────────────────────────────────────────────
    print(f"→ 定位总图题：{total_caption!r}，调整居中")
    W.find_text(total_caption)
    P.wait(0.3)
    P.hotkey('home')
    P.wait(0.1)
    P.hotkey('shift', 'end')
    P.wait(0.1)
    _set_center_align()

    print("→ 保存关闭")
    W.save_close()
    print("完成！")
