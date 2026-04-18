"""
Layer 1 — 原子操作。
每个函数只做一件事，无任何业务逻辑。
"""
import time
import pyautogui

pyautogui.FAILSAFE = True   # 鼠标移到左上角立即中止

DELAY = 0.25  # 默认按键后等待（秒）


def wait(t: float = DELAY) -> None:
    time.sleep(t)


def press(key: str, delay: float = DELAY) -> None:
    pyautogui.press(key)
    wait(delay)


def hotkey(*keys: str, delay: float = DELAY) -> None:
    pyautogui.hotkey(*keys)
    wait(delay)


def type_text(text: str, interval: float = 0.05) -> None:
    import pyperclip
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')
    wait()


def click(x: int, y: int, delay: float = DELAY) -> None:
    pyautogui.click(x, y)
    wait(delay)


def move(x: int, y: int) -> None:
    pyautogui.moveTo(x, y)


def right_click(x: int, y: int, delay: float = DELAY) -> None:
    pyautogui.rightClick(x, y)
    wait(delay)


def screenshot():
    return pyautogui.screenshot()


def locate_image_box(
    template_path: str,
    confidence: float = 0.8,
    region: tuple[int, int, int, int] | None = None,
) -> tuple[int, int, int, int]:
    """
    在屏幕（或 region 区域）内找到模板，返回 (left, top, width, height)。
    region 格式：(left, top, width, height)。
    """
    from PIL import Image
    template = Image.open(template_path)
    kwargs: dict = dict(confidence=confidence)
    if region:
        kwargs['region'] = region
    box = pyautogui.locateOnScreen(template, **kwargs)
    if box is None:
        raise RuntimeError(f"屏幕上未找到模板：{template_path}")
    return int(box.left), int(box.top), int(box.width), int(box.height)


def find_and_click_image(
    template_path: str,
    confidence: float = 0.8,
    move_only: bool = False,
    region: tuple[int, int, int, int] | None = None,
) -> tuple[int, int]:
    """
    在屏幕（或 region 区域）内找到模板图片，移动鼠标到中心，默认左键单击。
    move_only=True 时只移动不点击。
    region 格式：(left, top, width, height)。
    """
    left, top, w, h = locate_image_box(template_path, confidence=confidence, region=region)
    cx, cy = left + w // 2, top + h // 2
    pyautogui.moveTo(cx, cy)
    if not move_only:
        pyautogui.click(cx, cy)
    wait()
    return cx, cy
