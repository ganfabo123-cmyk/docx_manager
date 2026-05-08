from pydantic import BaseModel, Field
from typing import List


class ImageDescription(BaseModel):
    figure_type: str = Field(description="图片类型，如折线图/柱状图/框架图/示意图/实验截图/表格截图/其他")
    topic_summary: str = Field(description="一句话主题摘要，用于分组判断")
    main_content: str = Field(description="图片内容的完整描述（2-4句），用于段落定位")
    key_concepts: List[str] = Field(description="图中出现的技术关键词/方法名，用于精准匹配段落")
    suggested_caption: str = Field(description="推荐图题，格式为'图  简短描述'（图号留空）")


class ImageGroupingResponse(BaseModel):
    groups: List[List[int]] = Field(description="分组结果，每个子列表是一组图片的 index，所有图片必须被分配")


class ImageGroup(BaseModel):
    image_indices: List[int] = Field(description="组内图片在原始列表中的索引，按展示顺序排列")
    anchor_idx: int = Field(description="锚点段落在段落列表中的索引，图片组插入该段落之后")
    captions: List[str] = Field(description="组内每张图片对应的图题，与 image_indices 等长")


class HeadingSectionResponse(BaseModel):
    heading_id: str = Field(description="最适合放置该图片组的章节标题的 id 字符串")


class ImageGroupPlacement(BaseModel):
    anchor_idx: int = Field(description="图片组插入位置：从候选段落列表中选一个段落的 index 值，图片插入在该段落之后")
