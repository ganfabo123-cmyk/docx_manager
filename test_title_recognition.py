import json
from llm_data_collector.utils.generate_user_data import is_special_section_title

# 测试标题识别
test_titles = [
    "摘  要",  # 带空格
    "摘要",    # 不带空格
    "Abstract", 
    "结  论",  # 带空格
    "结论",    # 不带空格
    "致  谢",  # 带空格
    "致谢",    # 不带空格
    "已发表的学术论文目录",
    "附录A  OMML 参考速查",
    "参考文献"
]

print("测试标题识别:")
for title in test_titles:
    result = is_special_section_title(title)
    print(f"标题: '{title}' -> 识别结果: {result}")
