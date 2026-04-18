"""
Layer 3 — 图片排版工作流。

输入：
  docx_path  : 文档路径
  anchor_text: 定位段落文字
  image_path : 图片文件绝对路径
  caption    : 图片标题文字（图片下方居中段落）

环绕方式：嵌入型（WPS 插图默认，学术论文标准）
"""
from .. import wps_nav as W
from .. import primitives as P

_MARKER = "AAAAAAAA"


def _locate_caption_line(caption: str, offset_up: int = 40) -> tuple[int, int]:
    """
    OCR 扫描截图，定位 MARKER 所在行，上移点击后将 MARKER 替换为真正的图题文字。
    返回最终点击坐标。
    """
    P.hotkey("down")#光标下移,防止干扰
    
    screenshot = W.capture_screen()
    screenshot.save(r"D:\PycharmProjects\hit-paper-helper\ocr_debug.png")
    import numpy as np
    reader = W._get_ocr_reader()
    results = reader.readtext(np.array(screenshot), paragraph=False)
    print(f"   OCR 识别到 {len(results)} 条结果，搜索 {_MARKER!r}：")
    marker_center = None
    for bbox, text, conf in results:
        x1, y1 = int(bbox[0][0]), int(bbox[0][1])
        x2, y2 = int(bbox[2][0]), int(bbox[2][1])
        found = _MARKER in text
        tag = " <<<" if found else ""
        print(f"     ({x1:4d},{y1:4d})-({x2:4d},{y2:4d})  conf={conf:.3f}  {text!r}{tag}")
        if found and marker_center is None:
            marker_center = ((x1 + x2) // 2, (y1 + y2) // 2)

    if marker_center is None:
        raise RuntimeError(f"屏幕上未找到 MARKER {_MARKER!r}")

    tx, ty = marker_center[0], marker_center[1] - offset_up
    print(f"→ 定位图题行：MARKER 中心 {marker_center}，上移 {offset_up}px → 点击 ({tx}, {ty})")
    P.move(tx, ty)
    P.click(tx, ty)

    print(f"→ 替换 MARKER 为图题：{caption!r}")
    P.hotkey('up')#光标回归
    P.hotkey('ctrl', 'a')
    P.type_text(caption)

    return tx, ty


def open_image_property_panel() -> None:
    """右键菜单 → 属性面板：Shift+F10 → O"""
    P.hotkey('shift', 'f10')
    P.wait(0.4)
    P.press('o')
    P.wait(3)


def insert_image_after_paragraph(
    docx_path: str,
    anchor_text: str,
    image_path: str,
    caption: str,
) -> None:
    W.open_doc(docx_path)
    W.goto_start()

    print(f"→ 定位段落：{anchor_text!r}")
    W.find_text(anchor_text)
    W.goto_line_end()

    print("→ 插入图片")
    W.newline()
    W.newline()
    P.press('up')
    before_img = W.capture_screen()
    W.open_insert_picture_dialog()
    W.input_file_path_confirm(image_path)

    img_cx, _ = W.get_changed_center(before_img)
    print(f"   图片中心 X={img_cx}px")

    print(f"→ 插入标记占位符，OCR 定位")
    P.press('down')
    P.hotkey('ctrl', 'e')
    P.type_text(_MARKER)
    P.wait(0.5)

    _locate_caption_line(caption)
    print("调出属性选项卡")
    open_image_property_panel()
    print("寻找图片配置")
    W.navigate_to_crop_inputs()
    print("→ 保存关闭")
    W.save_close()
    print("完成！")
