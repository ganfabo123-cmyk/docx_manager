from unicodedata import name
from llm_data_collector.utils.parse_full_docx import parse_full_docx, parse_full_docx_simple

def parse_docx():
    # 返回完整结构（包含 headings 和 content）
    result = parse_full_docx("data\output.docx")
    print(result)
    # 只返回 content 内容
    content = parse_full_docx_simple("data\output.docx")
    #print(content)

if __name__ == "__main__":
    parse_docx()