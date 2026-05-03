from pydantic import BaseModel
from typing import List, Literal


class ShortBlockItem(BaseModel):
    id: str
    type: Literal["heading1", "heading2", "heading3", "normal"]


class ShortBlockListResponse(BaseModel):
    items: List[ShortBlockItem]
