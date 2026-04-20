from .insert_two_images import insert_two_images_after_paragraph
from .insert_image import insert_image_after_paragraph


def insert_n_images_two_col(docx_path: str, anchor_text: str, images: list[str], captions: list[str]):
    current_anchor = anchor_text
    i = 0
    while i < len(images):
        remaining = len(images) - i
        if remaining >= 2:
            insert_two_images_after_paragraph(
                docx_path=docx_path,
                anchor_text=current_anchor,
                image_path1=images[i],
                caption1=captions[i],
                image_path2=images[i + 1],
                caption2=captions[i + 1],
            )
            current_anchor = captions[i + 1]
            i += 2
        else:
            insert_image_after_paragraph(
                docx_path=docx_path,
                anchor_text=current_anchor,
                image_path=images[i],
                caption=captions[i],
            )
            current_anchor = captions[i]
            i += 1

def insert_n_images_one_col(docx_path: str, anchor_text: str, images: list[str], captions: list[str]):
    current_anchor = anchor_text
    i = 0
    while i < len(images):
        insert_image_after_paragraph(
                docx_path=docx_path,
                anchor_text=current_anchor,
                image_path=images[i],
                caption=captions[i],
            )
        current_anchor = captions[i]
        i += 1