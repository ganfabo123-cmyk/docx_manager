from pydantic import BaseModel, Field


class TableExtractResponse(BaseModel):
    is_not_table: bool = Field(
        description="若输入内容根本不是表格（如纯文字段落、公式、代码等），填 True；确认是表格数据则填 False。"
    )
    title: str = Field(
        description="表格标题。若原文在表格上方有标题行（如'表1 实验数据汇总'），提取到这里；找不到则填空字符串。"
    )
    content: list[list[str]] = Field(
        description=(
            "二维数组。第一行为列标题（保留单位信息，如'温度(℃)'）；"
            "其余每行对应一条数据记录，所有单元格均为字符串。"
            "去除 Markdown 加粗（**）、斜体（*）等格式符号，保留原始数值。"
            "is_not_table 为 True 时填空数组 []。"
        )
    )
