from unicodedata import name
from full_style_docx_fixer.utils.parse_full_docx import parse_full_docx
from full_style_docx_fixer.utils.generate_user_data import generate_user_data_from_file
def parse_docx():
    # 返回完整结构（包含 headings 和 content）
    result = parse_full_docx("docx_manager\data\\template.docx")
    parse_data = generate_user_data_from_file(docx_infos=result.get("docx_infos"))
    print(parse_data)


if __name__ == "__main__":
    parse_docx()