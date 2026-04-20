"""
最小 OCR 实现：截取 WPS 窗口，识别每行文本及其坐标。
依赖：pip install easyocr pywin32 pillow
模型路径：D:/ocr_models/
"""
# 跳过 EasyOCR 的 MD5 校验（本地模型被修改过时使用）
# easyocr.easyocr 直接 from .utils import calculate_md5，需 patch 该模块内的引用
import easyocr.easyocr as _ee
import easyocr.utils as _eu
_sentinel = object()
_orig_md5 = _eu.calculate_md5

def _skip_md5(path):
    """返回期望的 md5sum，使校验永远通过。"""
    return _sentinel

# 让比较式 calculate_md5(path) != model['md5sum'] 永远为 False
class _AlwaysEqual:
    def __eq__(self, other): return True
    def __ne__(self, other): return False

_ee.calculate_md5 = lambda path: _AlwaysEqual()

import easyocr
import numpy as np
from PIL import ImageGrab
import win32gui

MODEL_DIR = r"D:\ocr_models"


def get_wps_window():
    """找到 WPS 文档窗口句柄，返回 (left, top, right, bottom)。"""
    result = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "WPS" in title or ".docx" in title or ".doc" in title:
                result.append((hwnd, title))

    win32gui.EnumWindows(callback, None)
    if not result:
        raise RuntimeError("未找到 WPS 窗口，请先打开 WPS 文档")
    hwnd, title = result[0]
    print(f"找到窗口: {title}")
    return win32gui.GetWindowRect(hwnd)  # (left, top, right, bottom)


def capture_window(rect):
    img = ImageGrab.grab(bbox=rect)
    return np.array(img)


def ocr_lines(img_array, reader):
    """
    返回每行: {'text': str, 'bbox': [[x1,y1],...,[x1,y2]], 'conf': float}
    """
    results = reader.readtext(img_array, paragraph=False)
    lines = []
    for bbox, text, conf in results:
        lines.append({"text": text, "bbox": bbox, "conf": round(conf, 3)})
    return lines


def main():
    print("初始化 EasyOCR（首次稍慢）...")
    reader = easyocr.Reader(
        ["ch_sim", "en"],
        model_storage_directory=MODEL_DIR,
        download_enabled=True,
        gpu=True,   # 安装 CUDA 版 torch 后生效
    )

    rect = get_wps_window()
    print(f"窗口区域: {rect}")

    img = capture_window(rect)
    print(f"截图尺寸: {img.shape[1]}x{img.shape[0]}")

    lines = ocr_lines(img, reader)
    print(f"\n识别到 {len(lines)} 行文本:\n")
    for i, line in enumerate(lines):
        pts = line["bbox"]
        x1, y1 = int(pts[0][0]), int(pts[0][1])
        x2, y2 = int(pts[2][0]), int(pts[2][1])
        print(f"[{i+1:03d}] ({x1:4d},{y1:4d})-({x2:4d},{y2:4d})  conf={line['conf']}  {line['text']}")


if __name__ == "__main__":
    main()
