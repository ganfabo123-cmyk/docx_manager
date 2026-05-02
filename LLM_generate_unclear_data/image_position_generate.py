import json
from dataclasses import dataclass


@dataclass
class ImageGroup:
    image_indices: list[int]   # 本组图片在传入列表中的索引（从 0 开始）
    anchor_idx:    int         # 锚点段落在文档段落列表中的索引


def group_images(images: list, paragraphs: list[str]) -> list[list[int]]:
    from llm_router import route_group_images
    return route_group_images(images, paragraphs)


def sort_group_images(image_indices: list[int], paragraphs: list[str]) -> list[int]:
    from llm_router import route_sort_group_images
    return route_sort_group_images(image_indices, paragraphs)


def convert(position_json_str: str) -> list[ImageGroup]:
    data = json.loads(position_json_str)
    return [
        ImageGroup(
            image_indices=item["image_indices"],
            anchor_idx=item["anchor_idx"],
        )
        for item in data
    ]


def generate(images: list, paragraphs: list[str]) -> list[ImageGroup]:
    groups = group_images(images, paragraphs)
    sorted_groups = [sort_group_images(group, paragraphs) for group in groups]
    from llm_router import route_anchor_idx
    return convert(route_anchor_idx(sorted_groups, paragraphs))
