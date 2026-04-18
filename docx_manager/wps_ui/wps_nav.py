"""
Layer 2 — WPS 文档导航与对话框操作。
组合 primitives，封装 WPS 的快捷键序列。每个函数对应一个 WPS 可感知的操作。
"""
import os
from pathlib import Path
from . import primitives as P

_ASSETS = Path(__file__).parent / 'assets'


# ── 文档生命周期 ───────────────────────────────────────────────────────────

def open_doc(path: str, wait_sec: float = 4.0) -> None:
    os.startfile(path)
    P.wait(wait_sec)

def save() -> None:
    P.hotkey('ctrl', 's')
    P.wait(1.5)

def close() -> None:
    P.hotkey('alt', 'f4')

def save_close() -> None:
    save()
    close()


# ── 光标导航 ───────────────────────────────────────────────────────────────

def goto_start() -> None:
    P.hotkey('ctrl', 'home')

def goto_end() -> None:
    P.hotkey('ctrl', 'end')

def jump_next_section() -> None:
    """跳到下一节：Ctrl+G → Alt+T → Esc"""
    P.hotkey('ctrl', 'g')
    P.wait(0.5)
    P.hotkey('alt', 't')
    P.wait(0.3)
    P.press('escape')


# ── 页码对话框 ─────────────────────────────────────────────────────────────

def open_page_number_dialog() -> None:
    """Alt → P → N → U → N  打开页码格式对话框"""
    P.press('alt');  P.wait(0.2)
    P.press('p')
    P.press('n')
    P.press('u')
    P.press('n')
    P.wait(0.5)

def page_dialog_move(down: int = 0, up: int = 0) -> None:
    """在对话框中用方向键选择样式"""
    for _ in range(down): P.press('down')
    for _ in range(up):   P.press('up')

def page_dialog_apply_this_section_onward() -> None:
    """应用范围 = 本节及以后：Alt+P"""
    P.hotkey('alt', 'p')

def confirm() -> None:
    P.press('return')


# ── 文本查找 ───────────────────────────────────────────────────────────────

def find_text(text: str) -> None:
    """Ctrl+F → 输入文本 → Enter → Esc（光标停在匹配处）"""
    P.hotkey('ctrl', 'f')
    P.wait(0.5)
    P.type_text(text)
    P.press('return')
    P.wait(0.3)
    P.press('escape')

def goto_line_end() -> None:
    P.press('end')

def newline() -> None:
    P.press('return')


# ── 图片位置感知 ───────────────────────────────────────────────────────────

def capture_screen():
    return P.screenshot()


def get_changed_center(before) -> tuple[int, int]:
    """
    截图差分，返回变化最集中区域的中心 (cx, cy)。
    X：变化列范围的中点（图片横跨全宽，取中点合理）。
    Y：行变化量最大的那一行（argmax），抗滚动条/页面滚动噪声。
    """
    import numpy as np
    after = P.screenshot()
    diff = np.abs(
        np.array(after).astype(int) - np.array(before).astype(int)
    ).sum(axis=2)
    changed_cols = np.where(diff.max(axis=0) > 30)[0]
    if len(changed_cols) == 0:
        raise RuntimeError("未检测到变化区域")
    cx = int((changed_cols.min() + changed_cols.max()) // 2)
    cy = int(np.argmax(diff.sum(axis=1)))
    return cx, cy


# 保留旧名称兼容
get_image_center = get_changed_center


# ── 屏幕文字识别点击 ───────────────────────────────────────────────────────

_ocr_reader = None
OCR_MODEL_DIR: str | None = r'D:\ocr_models'

def _patch_easyocr_md5() -> None:
    """用本地文件的实际 MD5 覆盖 easyocr 注册表，跳过版本校验。"""
    import hashlib, os
    import easyocr.config as ec

    def actual_md5(filename: str) -> str | None:
        path = os.path.join(OCR_MODEL_DIR, filename)
        if not os.path.exists(path):
            return None
        h = hashlib.md5()
        with open(path, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                h.update(chunk)
        return h.hexdigest()

    for info in ec.detection_models.values():
        if isinstance(info, dict) and 'filename' in info:
            md5 = actual_md5(info['filename'])
            if md5:
                info['md5sum'] = md5

    for lang_models in ec.recognition_models.values():
        for info in lang_models.values():
            if isinstance(info, dict) and 'filename' in info:
                md5 = actual_md5(info['filename'])
                if md5:
                    info['md5sum'] = md5


def _get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        import easyocr
        if OCR_MODEL_DIR:
            _patch_easyocr_md5()
        kwargs: dict = dict(gpu=True, verbose=False)
        if OCR_MODEL_DIR:
            kwargs['model_storage_directory'] = OCR_MODEL_DIR
            kwargs['download_enabled'] = False
        _ocr_reader = easyocr.Reader(['ch_sim', 'en'], **kwargs)
    return _ocr_reader


def find_and_click_text(
    text: str,
    region: tuple[int, int, int, int] | None = None,
) -> tuple[int, int]:
    """
    截图 OCR 识别文字，点击第一个匹配项，返回点击坐标。
    region: (left, top, width, height)，None 为全屏。
    建议传入右侧面板区域以加快识别速度。
    """
    import numpy as np
    screenshot = P.screenshot()
    if region:
        left, top, w, h = region
        img = screenshot.crop((left, top, left + w, top + h))
        ox, oy = left, top
    else:
        img = screenshot
        ox, oy = 0, 0
    reader = _get_ocr_reader()
    for bbox, detected in reader.readtext(np.array(img), paragraph=True):
        if text in detected:
            cx = int((bbox[0][0] + bbox[2][0]) / 2) + ox
            cy = int((bbox[0][1] + bbox[2][1]) / 2) + oy
            P.click(cx, cy)
            return cx, cy
    raise RuntimeError(f"屏幕上未找到文字：{text!r}")


# ── 图片大小（OCR 方式）─────────────────────────────────────────────────────

def set_picture_size_via_panel(
    cx: int,
    cy: int,
    width_cm: float,
    height_cm: float,
    panel_region: tuple[int, int, int, int] | None = None,
) -> None:
    """
    通过右侧属性面板设置图片大小：
      左键点击图片 → Shift+F10 → O 打开属性面板
      → OCR 点击「图片」→「裁剪」→ 填宽度、高度。
    panel_region: 右侧面板屏幕区域 (left, top, width, height)，传入可加速 OCR。
    """
    P.click(cx, cy)
    P.wait(0.3)
    P.hotkey('shift', 'f10')
    P.wait(0.4)
    P.press('o')
    P.wait(0.6)
    find_and_click_text('图片', region=panel_region)
    P.wait(0.4)
    find_and_click_text('裁剪', region=panel_region)
    P.wait(0.4)
    find_and_click_text('宽度', region=panel_region)
    P.wait(0.2)
    P.press('tab')
    P.hotkey('ctrl', 'a')
    P.type_text(str(width_cm))
    P.press('return')
    P.wait(0.2)
    find_and_click_text('高度', region=panel_region)
    P.wait(0.2)
    P.press('tab')
    P.hotkey('ctrl', 'a')
    P.type_text(str(height_cm))
    P.press('return')


# ── 图片属性面板导航（模板匹配）─────────────────────────────────────────────

def navigate_to_crop_inputs(confidence: float = 0.9, input_offset_right: int = 60) -> tuple[tuple[int,int], tuple[int,int]]:
    """
    在已打开的属性面板中，依次点击：
      属性(包含图片) → 图片按钮 → 裁剪 → 宽度输入框 → 高度输入框
    每步找到元素后，将其所在列区域向下延伸作为后续搜索范围，避免误匹配其他图片的同名控件。
    返回 (宽度中心坐标, 高度中心坐标)，供后续输入使用。
    """
    import pyautogui as _pag

    # Step 1: 全屏找属性选项卡，点击并锁定面板区域供后续复用
    print("→ 模板定位：panel_tab_image")
    tab_left, tab_top, tab_w, tab_h = P.locate_image_box(
        str(_ASSETS / 'panel_tab_image.png'), confidence=confidence
    )
    screen_w, screen_h = _pag.size()
    panel_region = (tab_left, tab_top, screen_w - tab_left, screen_h - tab_top)
    _pag.click(tab_left + tab_w // 2, tab_top + tab_h // 2)
    P.wait(0.3)

    # Step 2-3: 在面板区域内依次点击
    for filename in ['panel_image_btn.png', 'panel_image_crop.png']:
        print(f"→ 模板定位：{filename}")
        P.find_and_click_image(str(_ASSETS / filename), confidence=confidence, region=panel_region)
        P.wait(0.3)

    print("→ 模板定位：宽度输入框")
    wx, wy = P.find_and_click_image(
        str(_ASSETS / 'panel_crop_width.png'), confidence=confidence, region=panel_region
    )
    width_pos = (wx + input_offset_right, wy)
    P.move(*width_pos)
    P.wait(0.2)

    print("→ 模板定位：高度输入框")
    hx, hy = P.find_and_click_image(
        str(_ASSETS / 'panel_crop_height.png'), confidence=confidence, region=panel_region
    )
    height_pos = (hx + input_offset_right, hy)
    P.move(*height_pos)

    return width_pos, height_pos


# ── 插入图片对话框 ─────────────────────────────────────────────────────────

def open_insert_picture_dialog() -> None:
    """Alt → N → P  打开插入图片对话框（WPS 插入选项卡）"""
    P.press('alt');  P.wait(0.2)
    P.press('n')
    P.press('p')
    P.press('p')
    P.wait(1.0)


def input_file_path_confirm(path: str) -> None:
    """在文件对话框地址栏粘贴路径并确认"""
    import pyperclip
    pyperclip.copy(path)
    P.hotkey('ctrl', 'l')
    P.wait(0.3)
    P.hotkey('ctrl', 'a')
    P.hotkey('ctrl', 'v')
    P.wait(0.3)
    P.press('return')
    P.wait(1.0)
    P.press('return')
