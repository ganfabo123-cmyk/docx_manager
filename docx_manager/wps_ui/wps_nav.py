"""
Layer 2 — WPS 文档导航与对话框操作。
组合 primitives，封装 WPS 的快捷键序列。每个函数对应一个 WPS 可感知的操作。
"""
import os
from pathlib import Path
from . import primitives as P

_ASSETS = Path(__file__).parent / 'assets'


# ── 文档生命周期 ───────────────────────────────────────────────────────────

def open_doc(path: str, wait_sec: float = 4) -> None:
    os.startfile(path)
    P.wait(wait_sec)
    maximize_window()

def save() -> None:
    P.hotkey('ctrl', 's')
    P.wait(0.1)

def close() -> None:
    P.hotkey('alt', 'f4')

def save_close(delay_before_close: float = 2.0) -> None:
    save()
    P.wait(delay_before_close)
    close()


# ── 光标导航 ───────────────────────────────────────────────────────────────

def goto_start() -> None:
    P.hotkey('ctrl', 'home')

def goto_end() -> None:
    P.hotkey('ctrl', 'end')

def jump_next_section() -> None:
    """跳到下一节：Ctrl+G → Alt+T → Esc"""
    P.hotkey('ctrl', 'g')
    P.wait(0.1)
    P.hotkey('alt', 't')
    P.wait(0.1)
    P.press('escape')


# ── 窗口管理 ───────────────────────────────────────────────────────────────

def maximize_window() -> None:
    """Win+↑ 最大化当前窗口，确保 ribbon 完整展开后再操作快捷键"""
    P.hotkey('win', 'up')
    P.wait(1.5)


# ── 页码对话框 ─────────────────────────────────────────────────────────────

def open_page_number_dialog() -> None:
    """Alt → P → N → U → N  打开页码格式对话框（需全屏，否则 N/U 会打入文档）"""
    maximize_window()
    P.press('alt');  P.wait(0.3)
    P.press('p')
    P.wait(0.3)
    P.press('n')
    P.wait(0.3)
    P.press('u')
    P.wait(0.3)
    P.press('n')
    P.wait(0.3)

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
    P.wait(0.1)
    P.type_text(text)
    P.press('return')
    P.wait(0.1)
    P.press('escape')

def goto_line_end() -> None:
    P.press('end')

def newline() -> None:
    P.press('return')


# ── 图片位置感知 ───────────────────────────────────────────────────────────

def capture_screen():
    return P.screenshot()


# ── 图片属性面板导航（模板匹配）─────────────────────────────────────────────
def navigate_to_crop_inputs(
    confidence: float = 0.8,
    img_width: float = None,
    img_height: float = None,
    crop_width: float = None,
    crop_height: float = None,
    offset_x: float = 0,
    offset_y: float = 0,
    click_crop: bool = True
) -> None:
    """
    在已打开的属性面板中，定位到图片 → 裁剪，然后使用 Alt+字母 快捷键修改各项参数。
    
    快捷键对照：
        W - 图片位置宽度
        H - 图片位置高度
        I - 裁剪位置宽度
        G - 裁剪位置高度
        X - 图片位置偏移 X（需按两次 Alt+X 才能激活）
        Y - 图片位置偏移 Y
    
    Args:
        confidence: 图片识别置信度
        img_width: 图片位置宽度值
        img_height: 图片位置高度值
        crop_width: 裁剪位置宽度值
        crop_height: 裁剪位置高度值
        offset_x: 图片位置偏移 X 值
        offset_y: 图片位置偏移 Y 值
        click_crop: 是否点击裁剪按钮
    """
    import pyautogui as _pag

    try:
        # Step 1: 全屏找属性选项卡，点击并锁定面板区域供后续复用
        print("→ 模板定位：panel_tab_image")
        tab_left, tab_top, tab_w, tab_h = P.locate_image_box(
            str(_ASSETS / 'panel_tab_image.png'), confidence=confidence
        )
        screen_w, screen_h = _pag.size()
        panel_region = (tab_left, tab_top, screen_w - tab_left, screen_h - tab_top)
        P.move(tab_left + tab_w // 2, tab_top + tab_h // 2)
        P.wait(1)

        # Step 2-3: 点击图片按钮和裁剪按钮
        steps = ['panel_image_btn.png', 'panel_image_crop.png'] if click_crop else ['panel_image_btn.png']
        for filename in steps:
            print(f"→ 模板定位：{filename}")
            P.find_and_click_image(str(_ASSETS / filename), confidence=confidence, region=panel_region)
            P.wait(1)

        # 确保焦点在属性面板内
        print("→ 准备使用快捷键修改参数...")
        P.wait(0.1)

        # 快捷键映射：参数名 -> (Alt+字母, 值, 是否需要双击)
        # 顺序：W H I G X Y
        shortcuts = [
            ('W', img_width, False),      # 图片位置宽度
            ('H', img_height, False),     # 图片位置高度
            ('I', crop_width, False),     # 裁剪位置宽度
            ('G', crop_height, False),    # 裁剪位置高度
            ('X', offset_x, True),        # 图片位置偏移 X（需要按两次）
            ('Y', offset_y, False),       # 图片位置偏移 Y
        ]

        for letter, value, need_double in shortcuts:
            if value is not None:
                print(f"→ Alt+{letter} 设置值为: {value}")
                
                if need_double:
                    # 偏移 X 需要按两次 Alt+X 才能激活编辑
                    _pag.hotkey('alt', letter.lower())
                    P.wait(0.1)
                    _pag.hotkey('alt', letter.lower())
                    P.wait(0.1)
                else:
                    _pag.hotkey('alt', letter.lower())
                    P.wait(0.1)
                
                # 全选并输入新值
                _pag.hotkey('ctrl', 'a')
                _pag.typewrite(str(value))
                # 按 Enter 确认，防止数据丢失
                _pag.press('enter')
                P.wait(0.1)

        print("→ 所有参数设置完成")
    except Exception as e:
        raise e


def set_image_size_by_shortcut(
    img_width: float = None,
    img_height: float = None,
    crop_width: float = None,
    crop_height: float = None,
    offset_x: float = 0,
    offset_y: float = 0,
) -> None:
    """
    属性面板已停在图片大小界面时，直接用 Alt 快捷键填值，无需 CV2 重新定位。
    快捷键同 navigate_to_crop_inputs：W/H=图片宽高，I/G=裁剪宽高，X/Y=偏移。
    """
    import pyautogui as _pag
    shortcuts = [
        ('w', img_width,   False),
        ('h', img_height,  False),
        ('i', crop_width,  False),
        ('g', crop_height, False),
        ('x', offset_x,    True),
        ('y', offset_y,    False),
    ]
    for letter, value, need_double in shortcuts:
        if value is not None:
            print(f"→ Alt+{letter.upper()} 设置值为: {value}")
            if need_double:
                _pag.hotkey('alt', letter)
                P.wait(0.1)
                _pag.hotkey('alt', letter)
            else:
                _pag.hotkey('alt', letter)
            P.wait(0.1)
            _pag.hotkey('ctrl', 'a')
            _pag.typewrite(str(value))
            _pag.press('enter')
            P.wait(0.1)


# ── 插入图片对话框 ─────────────────────────────────────────────────────────

def open_insert_picture_dialog() -> None:
    """Ctrl+M  打开插入图片对话框（自定义快捷键）"""
    P.hotkey('ctrl', 'm')
    P.wait(0.1)


def input_file_path_confirm(path: str) -> None:
    """在文件对话框地址栏粘贴路径并确认"""
    import pyperclip
    pyperclip.copy(path)
    P.hotkey('ctrl', 'l')
    P.wait(0.1)
    P.hotkey('ctrl', 'a')
    P.wait(0.1)
    P.hotkey('ctrl', 'v')
    P.wait(0.1)
    P.press('enter')
    P.wait(0.1)
    P.press('enter')