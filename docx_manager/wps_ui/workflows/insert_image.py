"""
Layer 3 — 图片排版工作流。

输入：
  docx_path  : 文档路径
  anchor_text: 定位段落文字
  image_path : 图片文件绝对路径
  caption    : 图片标题文字（图片下方居中段落）

环绕方式：嵌入型（WPS 插图默认，学术论文标准）
定位方式：插入后用 pyautogui 模板匹配在屏幕上找到图片中心，无需 OCR。
"""
from .. import wps_nav as W
from .. import primitives as P

HIT_DEFAULT_SINGLE_IMG_WIDTH = 12.00
HIT_DEFAULT_SINGLE_IMG_HEIGHT = 6.00


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


def insert_image_after_paragraph(
    docx_path: str,
    anchor_text: str,
    image_path: str,
    caption: str,
    close_after: bool = True,
    cnt: int = 0
) -> None:
    W.open_doc(docx_path)
    W.goto_start()
    P.wait(0.5)

    print(f"→ 定位段落：{anchor_text!r}")
    W.find_text(anchor_text)
    W.goto_line_end()
    P.wait(0.5)

    print("→ 插入图片")
    #W.newline()
    #W.newline()
    #P.press('up')
    P.hotkey('enter')
    P.wait(0.5)
    P.wait(0.5)
    W.open_insert_picture_dialog()
    W.input_file_path_confirm(image_path)
    P.wait(0.5)

    print(f"→ 写图题：{caption!r}")
    P.hotkey('right')
    P.wait(0.5)
    P.wait(0.5)
    P.hotkey('enter')
    P.wait(0.5)
   # P.hotkey('ctrl', 'e')
    P.type_text(caption)
    P.wait(1)
    P.hotkey('home')
    P.wait(1)
    P.hotkey('shift','end')
    P.wait(1)
    P.hotkey('alt')
    P.wait(0.5)
    P.hotkey('H')
    P.wait(0.5)
    P.hotkey('A')
    P.wait(0.5)
    P.hotkey('L')
    P.wait(0.5)
    P.hotkey('alt')
    P.wait(0.5)
    P.hotkey('H')
    P.wait(0.5)
    P.hotkey('A')
    P.wait(0.5)
    P.hotkey('C')
    P.wait(0.5)
    P.hotkey('home')
    P.wait(0.5)
    P.hotkey('shift','end')
    P.wait(0.5)
    P.hotkey('alt')
    P.wait(0.5)
    P.hotkey('o')
    P.wait(0.5)
    P.hotkey('p')
    P.wait(0.5)
    P.hotkey('y')
    P.wait(0.5)
    P.hotkey('0')
    P.wait(0.5)
    P.hotkey('enter')
    P.wait(0.5)
    P.wait(0.5)

    print("→ pyautogui 定位图片")
    img_cx, img_cy = _find_image_center(image_path,threshold=0.2)
    P.wait(0.5)

    print("→ 调出属性面板，设置图片尺寸")
    open_image_property_panel(img_cx, img_cy)
    
    click_crop = True
    if cnt > 0:
        click_crop = False
    W.navigate_to_crop_inputs(img_width=HIT_DEFAULT_SINGLE_IMG_WIDTH, img_height=HIT_DEFAULT_SINGLE_IMG_HEIGHT,crop_width=HIT_DEFAULT_SINGLE_IMG_WIDTH,crop_height=HIT_DEFAULT_SINGLE_IMG_HEIGHT,click_crop=click_crop)
    P.wait(0.5)
    P.click(img_cx, img_cy)
    P.wait(0.5)
    P.hotkey('left')
    P.wait(0.5)
    P.hotkey('backspace')
    P.wait(0.5)
    P.hotkey('backspace')
    P.wait(0.5)
    P.hotkey('alt')
    P.wait(0.5)
    P.hotkey('H')
    P.wait(0.5)
    P.hotkey('A')
    P.wait(0.5)
    P.hotkey('L')
    P.wait(0.5)
    P.hotkey('alt')
    P.wait(0.5)
    P.hotkey('H')
    P.wait(0.5)
    P.hotkey('A')
    P.wait(0.5)
    P.hotkey('C')
    P.wait(0.5)

    if close_after:
        print("→ 保存关闭")
        W.save_close()
        print("完成！")
    else:
        print("→ 仅保存")
        P.hotkey('ctrl', 's')
        P.wait(0.5)
        P.wait(0.5)
        print("完成！")


def insert_n_image_after_paragraph(
    docx_path: str,
    anchor_text: str,
    items: list[tuple[str, str]],
) -> None:
    """
    items: [(image_path, caption), ...]
    第一轮用 anchor_text 定位，后续每轮用上一轮的 caption 作为 anchor。
    """
    for i, (image_path, caption) in enumerate(items):
        anchor = anchor_text if i == 0 else items[i - 1][1]
        is_last = (i == len(items) - 1)
        print(f"→ 第 {i+1}/{len(items)} 张，anchor={anchor!r}")
        insert_image_after_paragraph(
            docx_path=docx_path,
            anchor_text=anchor,
            image_path=image_path,
            caption=caption,
            close_after=is_last,
            cnt=i
        )