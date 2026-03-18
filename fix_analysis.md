# 临时修复图片解析问题的脚本
import re

# 原始的图片解析代码问题分析
# 问题：嵌套层级不正确，使用了不存在的标签 "picFills"

# 修复方案：简化嵌套层级，直接从 blipFill 中查找 blip 元素

def parse_image_fixed(paragraph, doc):
    """
    修复后的图片解析函数
    """
    for run in paragraph.runs:
        for drawing in _iter_elements_by_tag(run._element, "drawing"):
            # 直接从 drawing 中查找 blipFill 元素
            for blip_fill in _iter_elements_by_tag(drawing, "blipFill"):
                # 从 blipFill 中查找 blip 元素
                for blip in _iter_elements_by_tag(blip_fill, "blip"):
                    embed_attr = blip.get(qn("r:embed"))
                    if embed_attr:
                        print("找到图片嵌入引用:", embed_attr)
                        return True
    return False

print("图片解析问题分析:")
print("1. 原始代码使用了不存在的标签 'picFills'")
print("2. 应该直接从 blipFill 中查找 blip 元素")
print("3. 修复方案：简化嵌套层级")