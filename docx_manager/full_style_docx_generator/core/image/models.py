from pydantic import BaseModel, Field
from typing import List


class ImageGroup(BaseModel):
    image_indices: List[int] = Field(description="组内图片在原始列表中的索引，按展示顺序排列")
    anchor_idx: int = Field(description="锚点段落在段落列表中的索引，图片组插入该段落之后")
    captions: List[str] = Field(description="组内每张图片对应的图题，与 image_indices 等长")


class ImageGroupListResponse(BaseModel):
    groups: List[ImageGroup]
