import re
from typing import Optional
from docx.enum.text import WD_ALIGN_PARAGRAPH

# XML 命名空间
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

TAG_RE  = re.compile(r"^\[\[(.+?)\]\]$")
CITE_RE = re.compile(r"\[(\d+)\]")

HEADING_TAGS = {"一级标题", "二级标题", "三级标题"}
_STATIC_TAGS = HEADING_TAGS | {"正文", "表格", "参考文献", "公式", "图片", "目录"}

TYPE_TO_TAG = {
    "heading1": "一级标题",
    "heading2": "二级标题",
    "heading3": "三级标题",
}
# TAG_TO_STYLE_ID 保留作为模板里没有对应块时的最终兜底
TAG_TO_STYLE_ID = {
    "一级标题": "Heading1",
    "二级标题": "Heading2",
    "三级标题": "Heading3",
}

_ALIGN_MAP = {
    "left":   WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right":  WD_ALIGN_PARAGRAPH.RIGHT,
}

SECTION_DEFAULT_TITLES = {
    "abstract":        "摘  要",
    "abstract_en":     "Abstract",
    "conclusion":      "结  论",
    "acknowledgement": "致  谢",
    "publications":    "已发表的学术论文目录",
    "custom":          "",
}

# 页脚预设样式
FOOTER_STYLES: dict[str, Optional[dict]] = {
    "roman_lower_center": {
        "fmt": "lowerRoman", "prefix": "",   "suffix": "",
        "instr": "PAGE",     "align": "center", "with_numpages": False,
    },
    "roman_upper_center": {
        "fmt": "upperRoman", "prefix": "",   "suffix": "",
        "instr": "PAGE",     "align": "center", "with_numpages": False,
    },
    "arabic_center": {
        "fmt": "decimal",    "prefix": "",   "suffix": "",
        "instr": "PAGE",     "align": "center", "with_numpages": False,
    },
    "arabic_dash": {
        "fmt": "decimal",    "prefix": "- ", "suffix": " -",
        "instr": "PAGE",     "align": "center", "with_numpages": False,
    },
    "arabic_page_x": {
        "fmt": "decimal",    "prefix": "第", "suffix": "页",
        "instr": "PAGE",     "align": "center", "with_numpages": False,
    },
    "arabic_slash": {
        "fmt": "decimal",    "prefix": "",   "suffix": "",
        "instr": "PAGE",     "align": "center", "with_numpages": True,
    },
    "none": None,
}

# LLM 系统提示（页脚解析）
_FOOTER_LLM_SYSTEM = """你是一个 Word 文档格式助手。
用户会用自然语言描述文档的页码需求。你需要把它映射到预设样式配置。

可用的 section 值：
  frontmatter  → 正文（第一个一级标题）出现之前的所有页
  mainmatter   → 正文开始（第一个一级标题）到文档末尾

可用的 style 值：
  roman_lower_center   小写罗马数字居中 (i, ii, iii...)
  roman_upper_center   大写罗马数字居中 (I, II, III...)
  arabic_center        阿拉伯数字居中 (1, 2, 3...)
  arabic_dash          -X- 格式居中 (-1-, -2-...)
  arabic_page_x        第X页格式居中 (第1页, 第2页...)
  arabic_slash         X/N 格式居中 (1/10, 2/10...)
  none                 无页脚

输出严格的 JSON 数组，不含任何解释：
[
  {"section": "frontmatter", "style": "roman_lower_center", "start": 1},
  {"section": "mainmatter",  "style": "arabic_dash",        "start": 1}
]

如果整个文档只有一种页码，section 填 mainmatter，frontmatter 用 none。
只输出 JSON，不要 markdown 代码块，不要任何其他文字。"""
