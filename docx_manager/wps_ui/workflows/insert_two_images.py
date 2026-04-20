"""
Layer 3 — 双图横排排版工作流。

输入：
  docx_path   : 文档路径
  anchor_text : 定位段落文字
  image_path1 : 第一张图片绝对路径
  caption1    : 第一张图题
  image_path2 : 第二张图片绝对路径
  caption2    : 第二张图题
  caption_gap : 两图题之间的半角空格数（默认 8）

图题排布：两图题同行，中间以 caption_gap 个半角空格分隔。
定位方式：插入后用 OpenCV 模板匹配在屏幕上找到图片中心，无需占位 MARKER。
"""
from .. import wps_nav as W
from .. import primitives as P

HIT_DEFAULT_DOUBLE_IMG_WIDTH  = 6.99
HIT_DEFAULT_DOUBLE_IMG_HEIGHT = 4.99


def _find_image_center(image_path: str, threshold: float = 0.8) -> tuple[int, int]:
    import cv2
    import numpy as np
    import mss

    with mss.mss() as sct:
        monitor = sct.monitors[1]
        screen = np.array(sct.grab(monitor))
    screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGRA2GRAY)

    template = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    if template is None:
        raise RuntimeError(f"无法读取模板图片：{image_path}")

    best_val, best_loc, best_scale = 0, None, 1.0
    for scale in np.linspace(0.7, 1.3, 13):
        h, w = template.shape
        resized = cv2.resize(template, (max(1, int(w * scale)), max(1, int(h * scale))))
        result = cv2.matchTemplate(screen_gray, resized, cv2.TM_CCOEFF_NORMED)
        _, val, _, loc = cv2.minMaxLoc(result)
        if val > best_val:
            best_val, best_loc, best_scale = val, loc, scale

    if best_val < threshold:
        raise RuntimeError(f"屏幕上未找到图片（best={best_val:.3f}）：{image_path}")

    th = int(template.shape[0] * best_scale)
    tw = int(template.shape[1] * best_scale)
    cx = best_loc[0] + tw // 2
    cy = best_loc[1] + th // 2
    print(f"   CV2 定位成功，图片中心 ({cx}, {cy})，置信度 {best_val:.3f}，缩放 {best_scale:.2f}")
    return cx, cy


def open_image_property_panel(x: int, y: int) -> None:
    """右键点击图片 → 属性面板：右键 → O"""
    P.right_click(x, y)
    P.wait(0.4)
    P.press('o')
    P.wait(3)


def insert_two_images_after_paragraph(
    docx_path: str,
    anchor_text: str,
    image_path1: str,
    caption1: str,
    image_path2: str,
    caption2: str,
    caption_gap: int = 8,
) -> None:
    W.open_doc(docx_path)
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

    # ── 图1 插入 + 图题 ────────────────────────────────────────────────────
    print("→ 插入图1")
    W.open_insert_picture_dialog()
    W.input_file_path_confirm(image_path1)
    P.wait(0.5)

    print(f"→ 写图1图题：{caption1!r}")
    P.press('down')
    P.hotkey('ctrl', 'e')
    P.type_text(caption1)
    P.wait(0.5)

    print("→ OpenCV 定位图1")
    img1_cx, img1_cy = _find_image_center(image_path1)
    P.wait(0.5)

    print("→ 调出属性面板，设置图1尺寸")
    open_image_property_panel(img1_cx, img1_cy)
    W.navigate_to_crop_inputs(img_width=HIT_DEFAULT_DOUBLE_IMG_WIDTH,
                               img_height=HIT_DEFAULT_DOUBLE_IMG_HEIGHT)
    P.wait(0.5)
    P.click(img1_cx,img1_cy)
    P.hotkey('left')
    P.hotkey('backspace')
    P.hotkey('ctrl','e')
    P.wait(0.5)
    # ── 图2 插入 + 图题（同行） ────────────────────────────────────────────
    print("→ 回到图1行尾，插入图2")
    P.click(img1_cx, img1_cy)
    P.press('right')
    P.wait(0.5)

    print("→ 插入图2")
    W.open_insert_picture_dialog()
    W.input_file_path_confirm(image_path2)
    P.wait(0.5)

    print(f"→ 写图2图题（追加到同行）：{caption2!r}")
    P.press('down')
    P.press('end')
    P.type_text(' ' * caption_gap)
    P.type_text(caption2)
    P.wait(0.5)

    print("→ OpenCV 定位图2")
    img2_cx, img2_cy = _find_image_center(image_path2)
    P.wait(0.5)

    print("→ 调出属性面板，设置图2尺寸")
    open_image_property_panel(img2_cx, img2_cy)
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
    print("→ 保存关闭")
    W.save_close()
    print("完成！")
