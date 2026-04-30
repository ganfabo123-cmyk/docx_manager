"""
Layer 1 — 原子操作。
每个函数只做一件事，无任何业务逻辑。
"""
import time
import pyautogui
import cv2
import numpy as np
import mss
from PIL import Image

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


def _find_image_center_cv2(
    image_path: str,
    confidence: float = 0.8,
    region: tuple[int, int, int, int] | None = None
) -> tuple[int, int, int, int]:
    """
    使用 OpenCV + MSS 在屏幕上查找图片，支持多尺度匹配。
    返回 (center_x, center_y, width, height)
    
    Args:
        image_path: 模板图片路径
        confidence: 匹配置信度阈值 (0.0 ~ 1.0)
        region: 搜索区域 (left, top, width, height)，None 表示全屏搜索
    """
    with mss.mss() as sct:
        if region:
            # 如果有 region，只截取指定区域
            monitor = {
                "top": region[1],
                "left": region[0],
                "width": region[2],
                "height": region[3]
            }
        else:
            monitor = sct.monitors[1]
        screen = np.array(sct.grab(monitor))
    
    screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGRA2GRAY)

    template = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    if template is None:
        raise RuntimeError(f"无法读取模板图片：{image_path}")

    best_val, best_loc, best_scale = 0, None, 1.0
    best_w, best_h = template.shape[1], template.shape[0]
    
    # 多尺度匹配
    for scale in np.linspace(0.7, 1.3, 13):
        h, w = template.shape
        resized = cv2.resize(template, (max(1, int(w * scale)), max(1, int(h * scale))))
        if resized.shape[0] > screen_gray.shape[0] or resized.shape[1] > screen_gray.shape[1]:
            continue  # 缩放后比屏幕还大，跳过
            
        result = cv2.matchTemplate(screen_gray, resized, cv2.TM_CCOEFF_NORMED)
        _, val, _, loc = cv2.minMaxLoc(result)
        if val > best_val:
            best_val = val
            best_loc = loc
            best_scale = scale
            best_w, best_h = resized.shape[1], resized.shape[0]

    if best_val < confidence:
        raise RuntimeError(f"屏幕上未找到模板（best={best_val:.3f}）：{image_path}")

    # 计算中心点坐标
    if region:
        cx = region[0] + best_loc[0] + best_w // 2
        cy = region[1] + best_loc[1] + best_h // 2
        left = region[0] + best_loc[0]
        top = region[1] + best_loc[1]
    else:
        cx = best_loc[0] + best_w // 2
        cy = best_loc[1] + best_h // 2
        left = best_loc[0]
        top = best_loc[1]
    
    print(f"   CV2 定位成功: {image_path} -> 中心({cx}, {cy}), 置信度{best_val:.3f}, 缩放{best_scale:.2f}")
    return cx, cy, best_w, best_h


def locate_image_box(
    template_path: str,
    confidence: float = 0.8,
    region: tuple[int, int, int, int] | None = None,
) -> tuple[int, int, int, int]:
    """
    在屏幕（或 region 区域）内找到模板，返回 (left, top, width, height)。
    region 格式：(left, top, width, height)。
    """
    cx, cy, w, h = _find_image_center_cv2(template_path, confidence, region)
    left = cx - w // 2
    top = cy - h // 2
    return left, top, w, h


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
    cx, cy, w, h = _find_image_center_cv2(template_path, confidence, region)
    pyautogui.moveTo(cx, cy)
    wait(1)
    if not move_only:
        pyautogui.click(cx, cy)
    wait()
    return cx, cy