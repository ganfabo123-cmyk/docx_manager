"""
Layer 3 — 双列多图排版工作流。                                                                                                                     
                                                                                                                                                     输入（insert_n_images_two_col）：
  docx_path   : 文档路径                                                                                                                           
  anchor_text : 首对定位段落文字
  items       : [(image_path, caption), ...]，数量必须为偶数
  caption_gap : 同行两图题之间的半角空格数（默认 8）
图题排布：每对两图题同行，中间以 caption_gap 个半角空格分隔。
定位方式：插入后用 pyautogui 模板匹配在屏幕上找到图片中心，无需占位 MARKER。
"""
from .. import wps_nav as W
from .. import primitives as P
HIT_DEFAULT_DOUBLE_IMG_WIDTH  = 6.99
HIT_DEFAULT_DOUBLE_IMG_HEIGHT = 4.99
def _find_image_center(image_path: str, confidence: float = 0.09, grayscale: bool = True) -> tuple[int, int]:
    import pyautogui
    box = pyautogui.locateOnScreen(image_path, confidence=confidence, grayscale=grayscale)
    if box is None:
        raise RuntimeError(f"屏幕上未找到图片（confidence={confidence}）：{image_path}")
    cx = box.left + box.width // 2
    cy = box.top + box.height // 2
    print(f"   pyautogui 定位成功，图片中心 ({cx}, {cy})")
    return cx, cy
def open_image_property_panel(x: int, y: int) -> None:
    """右键点击图片 → 属性面板：右键 → O"""
    P.right_click(x, y)
    P.wait(0.4)
    P.press('o')
    P.wait(3)
def _insert_pair(
    anchor_text: str,
    left_img: str,
    left_cap: str,
    right_img: str,
    right_cap: str,
    caption_gap: int,
    close_after: bool,
    overall_left_idx: int,
) -> None:
    """
    在已打开的文档中插入一对横排图片。
    overall_left_idx=0 时首次打开属性面板需点击裁剪 tab，之后跳过。
    """
    W.goto_start()
    P.wait(0.5)
    print(f"→ 定位段落：{anchor_text!r}")
    W.find_text(anchor_text)
    W.goto_line_end()
    P.wait(0.5)
    W.newline()
    W.newline()
    P.press('up')
    P.wait(0.5)
    # ── 左图 插入 + 图题 ───────────────────────────────────────────────────
    print("→ 插入左图")
    W.open_insert_picture_dialog()
    W.input_file_path_confirm(left_img)
    P.wait(0.5)
    print(f"→ 写左图图题：{left_cap!r}")
    P.press('down')
    P.hotkey('ctrl', 'e')
    P.type_text(left_cap)
    P.wait(0.5)
    print("→ 定位左图")
    lx, ly = _find_image_center(left_img)
    P.wait(0.5)
    print("→ 调出属性面板，设置左图尺寸")
    open_image_property_panel(lx, ly)
    click_crop = (overall_left_idx == 0)
    W.navigate_to_crop_inputs(img_width=HIT_DEFAULT_DOUBLE_IMG_WIDTH,
                               img_height=HIT_DEFAULT_DOUBLE_IMG_HEIGHT,
                               click_crop=click_crop)
    P.wait(0.5)
    P.click(lx, ly)
    P.hotkey('left')
    P.hotkey('backspace')
    P.hotkey('ctrl', 'e')
    P.wait(0.5)
    # ── 右图 插入 + 图题（同行） ───────────────────────────────────────────
    print("→ 回到左图行尾，插入右图")
    P.click(lx, ly)
    P.press('right')
    P.wait(0.5)
    print("→ 插入右图")
    W.open_insert_picture_dialog()
    W.input_file_path_confirm(right_img)
    P.wait(0.5)
    print(f"→ 写右图图题（追加同行）：{right_cap!r}")
    P.press('down')
    P.press('end')
    P.type_text(' ' * caption_gap)
    P.type_text(right_cap)
    P.wait(0.5)
    print("→ 定位右图")
    rx, ry = _find_image_center(right_img)
    P.wait(0.5)
    print("→ 调出属性面板，设置右图尺寸")
    open_image_property_panel(rx, ry)
    W.navigate_to_crop_inputs(img_width=HIT_DEFAULT_DOUBLE_IMG_WIDTH,
                               img_height=HIT_DEFAULT_DOUBLE_IMG_HEIGHT,
                               click_crop=False)
    P.wait(0.5)
    P.hotkey('enter')
    P.hotkey('left')
    P.hotkey('backspace')
    P.hotkey('backspace')
    P.hotkey('space')
    P.hotkey('space')
    P.hotkey('space')
    if close_after:
        print("→ 保存关闭")
        W.save_close()
        print("完成！")
    else:
        print("→ 仅保存")
        P.hotkey('ctrl', 's')
        P.wait(0.5)
def insert_n_images_two_col(
    docx_path: str,
    anchor_text: str,
    items: list[tuple[str, str]],
    caption_gap: int = 8,
) -> None:
    """
    items: [(image_path, caption), ...]，数量必须为偶数。
    按对插入两列排版，第 2 对起以上一对右图 caption 为 anchor。
    """
    assert len(items) % 2 == 0, f"items 数量必须为偶数，当前为 {len(items)}"
    W.open_doc(docx_path)
    P.wait(0.5)
    pairs = [(items[i], items[i + 1]) for i in range(0, len(items), 2)]
    for pair_idx, ((left_img, left_cap), (right_img, right_cap)) in enumerate(pairs):
        anchor = anchor_text if pair_idx == 0 else pairs[pair_idx - 1][1][1]
        is_last = (pair_idx == len(pairs) - 1)
        print(f"→ 第 {pair_idx + 1}/{len(pairs)} 对，anchor={anchor!r}")
        _insert_pair(
            anchor_text=anchor,
            left_img=left_img,
            left_cap=left_cap,
            right_img=right_img,
            right_cap=right_cap,
            caption_gap=caption_gap,
            close_after=is_last,
            overall_left_idx=pair_idx * 2,
        )
def insert_two_images_after_paragraph(
    docx_path: str,
    anchor_text: str,
    image_path1: str,
    caption1: str,
    image_path2: str,
    caption2: str,
    caption_gap: int = 8,
) -> None:
    """兼容旧接口，内部调用 insert_n_images_two_col。"""
    insert_n_images_two_col(
        docx_path=docx_path,
        anchor_text=anchor_text,
        items=[(image_path1, caption1), (image_path2, caption2)],
        caption_gap=caption_gap,
    )