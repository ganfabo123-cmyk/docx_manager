"""
一次性标定常量。固定分辨率/WPS窗口/文档缩放后测量一次即可。

测量方法：
  TEXT_LEFT_X   — 打开文档截图，在 Paint 里量文字区域左边界的屏幕 X 像素值
  SPACE_WIDTH_PX — 在 WPS 空行输入 20 个半角空格并截图，量总宽度 ÷ 20
  CHAR_WIDTH_PX  — 用标题样本文字截图，量总像素宽 ÷ 字符数（中英混排取均值）
"""

TEXT_LEFT_X: int    = 0    # 文本区左边界屏幕 X（待填入）
SPACE_WIDTH_PX: float = 1.0  # 半角空格像素宽（待填入）
CHAR_WIDTH_PX: float  = 1.0  # 标题字符平均像素宽（待填入）
