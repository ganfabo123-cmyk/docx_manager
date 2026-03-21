from dataclasses import dataclass, field
from typing import Optional

@dataclass
class TocLevelStyle:
    """单级目录条目的样式（pPr + rPr）"""
    pPr: object = None   # lxml Element
    rPr: object = None   # lxml Element


@dataclass
class HeadingStyle:
    """单级标题的完整样式（pPr + rPr）"""
    pPr: object = None
    rPr: object = None


@dataclass
class StyleTemplate:
    """从模板 docx 中提取的各类格式样板。所有格式字段都来自模板，None 表示模板未提供。"""
    table_proto:          object = None   # 表格 lxml Element
    table_caption_pPr:    object = None   # 表格标题段落格式
    table_caption_rPr:    object = None   # 表格标题文字格式
    ref_pPr_proto:        object = None   # 参考文献条目段落格式
    ref_rPr_proto:        object = None   # 参考文献条目文字格式
    formula_pPr:          object = None   # 公式段落格式
    formula_label_rPr:    object = None   # 式(x) 编号文字格式
    image_pPr:            object = None   # 图片段落格式（图片行）
    image_caption_pPr:    object = None   # 图片 caption 段落格式
    caption_rPr:          object = None   # 图片 caption 文字格式
    section_styles:       dict   = field(default_factory=dict)

    # 标题样式：{"heading1": HeadingStyle, "heading2": ..., "heading3": ...}
    heading_styles:       dict   = field(default_factory=dict)

    # 正文样式
    body_pPr:             object = None
    body_rPr:             object = None

    # 目录各级样式
    toc_level_styles:     list   = field(default_factory=list)