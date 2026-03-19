import json
from llm_data_collector.utils.generate_user_data import generate_user_data, load_config

# 模拟 parse_full_docx 生成的数据
test_data = [
    {
        "type": "heading1",
        "value": "摘  要",
        "chunk_id": "heading1_1",
        "between": ["START", "body_2"]
    },
    {
        "type": "body",
        "value": "本文对 docx_inline_tag_engine v5 的全部功能进行了系统测试，涵盖前置章节、目录、标题、正文、表格、公式与图片等所有内容类型。",
        "chunk_id": "body_2",
        "between": ["heading1_1", "body_3"]
    },
    {
        "type": "body",
        "value": "测试数据包含三种公式输入方式（OMML 直通、LaTeX 转换、退化纯文本）、四张图片（三种 base64 图、一张本地路径图）、三个不同结构的表格，以及 toc_exclude 排除机制的正反两向测试。",
        "chunk_id": "body_3",
        "between": ["body_2", "body_4"]
    },
    {
        "type": "body",
        "value": "关键词：文档自动化；排版引擎；全功能测试；Word XML",
        "chunk_id": "body_4",
        "between": ["body_3", "heading1_5"]
    },
    {
        "type": "heading1",
        "value": "Abstract",
        "chunk_id": "heading1_5",
        "between": ["body_4", "body_6"]
    },
    {
        "type": "body",
        "value": "This document provides a comprehensive test of all features in docx_inline_tag_engine v5, covering front matter sections, table of contents, headings at three levels, body text, tables, formulas, and images.",
        "chunk_id": "body_6",
        "between": ["heading1_5", "body_7"]
    },
    {
        "type": "body",
        "value": "Test cases include three formula input modes (OMML pass-through, LaTeX conversion, and plain-text fallback), four images (three base64-encoded, one local-path), three tables of varying structure, and bidirectional tests of the toc_exclude mechanism.",
        "chunk_id": "body_7",
        "between": ["body_6", "body_8"]
    },
    {
        "type": "body",
        "value": "Keywords: document automation; typesetting engine; full-feature test; Word XML",
        "chunk_id": "body_8",
        "between": ["body_7", "heading1_9"]
    },
    {
        "type": "heading1",
        "value": "第一章  引言",
        "chunk_id": "heading1_9",
        "between": ["body_8", "body_10"]
    },
    {
        "type": "body",
        "value": "docx_inline_tag_engine 是一套基于模板驱动的 Word 文档自动排版引擎。",
        "chunk_id": "body_10",
        "between": ["heading1_9", "heading1_11"]
    },
    {
        "type": "heading1",
        "value": "结  论",
        "chunk_id": "heading1_11",
        "between": ["body_10", "body_12"]
    },
    {
        "type": "body",
        "value": "本测试文档验证了 docx_inline_tag_engine v5 的全部功能模块：",
        "chunk_id": "body_12",
        "between": ["heading1_11", "body_13"]
    },
    {
        "type": "body",
        "value": "（1）前置章节（摘要/英文摘要）：标题样式独立，toc_exclude 正常工作。",
        "chunk_id": "body_13",
        "between": ["body_12", "body_14"]
    },
    {
        "type": "body",
        "value": "（2）目录：manual 模式支持前置页罗马数字 + 正文阿拉伯数字混用。",
        "chunk_id": "body_14",
        "between": ["body_13", "heading1_15"]
    },
    {
        "type": "heading1",
        "value": "致  谢",
        "chunk_id": "heading1_15",
        "between": ["body_14", "body_16"]
    },
    {
        "type": "body",
        "value": "感谢 python-docx 开源社区提供了优秀的 Word XML 操作接口。",
        "chunk_id": "body_16",
        "between": ["heading1_15", "body_17"]
    },
    {
        "type": "body",
        "value": "感谢所有提供反馈和建议的用户，你们的意见直接推动了引擎功能的完善。",
        "chunk_id": "body_17",
        "between": ["body_16", "heading1_18"]
    },
    {
        "type": "heading1",
        "value": "已发表的学术论文目录",
        "chunk_id": "heading1_18",
        "between": ["body_17", "body_19"]
    },
    {
        "type": "body",
        "value": "1. 张三, 李四. 基于标记驱动的 Word 文档自动排版引擎[J]. 计算机应用研究, 2025, 42(6): 1-8.",
        "chunk_id": "body_19",
        "between": ["heading1_18", "body_20"]
    },
    {
        "type": "body",
        "value": "2. Zhang S, Li S. Tag-driven Automatic Typesetting for Academic Documents[C]. ICDAR 2025, 2025: 234-241.",
        "chunk_id": "body_20",
        "between": ["body_19", "heading1_21"]
    },
    {
        "type": "heading1",
        "value": "附录A  OMML 参考速查",
        "chunk_id": "heading1_21",
        "between": ["body_20", "body_22"]
    },
    {
        "type": "body",
        "value": "常用 OMML 元素速查：",
        "chunk_id": "body_22",
        "between": ["heading1_21", "body_23"]
    },
    {
        "type": "body",
        "value": "m:f        → 分数（num/den）",
        "chunk_id": "body_23",
        "between": ["body_22", "body_24"]
    },
    {
        "type": "body",
        "value": "m:sSup     → 上标（e/sup）",
        "chunk_id": "body_24",
        "between": ["body_23", "heading1_25"]
    },
    {
        "type": "heading1",
        "value": "参考文献",
        "chunk_id": "heading1_25",
        "between": ["body_24", "body_26"]
    },
    {
        "type": "body",
        "value": "[1] Knuth D E. The TeXbook[M]. Reading, MA: Addison-Wesley Professional, 1984.",
        "chunk_id": "body_26",
        "between": ["heading1_25", "body_27"]
    },
    {
        "type": "body",
        "value": "[2] Clark A. python-docx: Create and update Microsoft Word .docx files[EB/OL]. https://python-docx.readthedocs.io, 2023.",
        "chunk_id": "body_27",
        "between": ["body_26", "END"]
    }
]

# 加载配置
config = load_config('llm_data_collector/utils/user_config.json')

# 测试转换
result = generate_user_data(test_data, config)

# 保存测试结果
with open('test_conversion_result.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print('测试完成，结果已保存到 test_conversion_result.json')
print('\n关键转换结果:')
print('\n1. 摘要转换:')
for item in result['content']:
    if item.get('section_type') == 'abstract':
        print(json.dumps(item, ensure_ascii=False, indent=2))

print('\n2. Abstract 转换:')
for item in result['content']:
    if item.get('section_type') == 'abstract_en':
        print(json.dumps(item, ensure_ascii=False, indent=2))

print('\n3. 结论转换:')
for item in result['content']:
    if item.get('section_type') == 'conclusion':
        print(json.dumps(item, ensure_ascii=False, indent=2))

print('\n4. 致谢转换:')
for item in result['content']:
    if item.get('section_type') == 'acknowledgement':
        print(json.dumps(item, ensure_ascii=False, indent=2))

print('\n5. 已发表的学术论文目录转换:')
for item in result['content']:
    if item.get('section_type') == 'publications':
        print(json.dumps(item, ensure_ascii=False, indent=2))

print('\n6. 附录转换:')
for item in result['content']:
    if item.get('section_type') == 'custom':
        print(json.dumps(item, ensure_ascii=False, indent=2))

print('\n7. 参考文献转换:')
if 'references' in result:
    print(json.dumps(result['references'], ensure_ascii=False, indent=2))
