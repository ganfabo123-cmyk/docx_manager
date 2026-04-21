import argparse
import json
from docx_manager.wps_ui.workflows.hit_footer import apply_hit_page_numbers
from docx_manager.wps_ui.workflows.insert_image import insert_n_image_after_paragraph


def _parse_item(raw: list[str]) -> tuple:
    """['path', 'caption'] 或 ['path', 'caption', 'w', 'h'] → tuple"""
    if len(raw) < 2 or len(raw) > 4:
        raise argparse.ArgumentTypeError(
            f"--item 需要 2~4 个值（path caption [width height]），收到：{raw}"
        )
    path, caption = raw[0], raw[1]
    width = float(raw[2]) if len(raw) > 2 and raw[2] else None
    height = float(raw[3]) if len(raw) > 3 and raw[3] else None
    return (path, caption, width, height)


def _load_items_from_json(json_path: str) -> tuple[str, list[tuple]]:
    """
    读取 JSON 文件，返回 (anchor_text, items)。

    JSON 格式：
    {
        "anchor_text": "第一章",
        "items": [
            {"path": "img1.png", "caption": "图1"},
            {"path": "img2.png", "caption": "图2", "width": 12.0, "height": 6.0}
        ]
    }
    """
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    anchor_text = data["anchor_text"]
    items = []
    for entry in data["items"]:
        items.append((
            entry["path"],
            entry["caption"],
            entry.get("width") or None,
            entry.get("height") or None,
        ))
    return anchor_text, items


def cmd_footer(args):
    apply_hit_page_numbers(args.docx_path, args.body_section)


def cmd_insert_image(args):
    if args.from_json:
        anchor_text, items = _load_items_from_json(args.from_json)
    elif args.item:
        anchor_text = args.anchor_text
        items = [_parse_item(raw) for raw in args.item]
    else:
        raise SystemExit("请通过 --from-json 或 --item 提供图片信息")

    insert_n_image_after_paragraph(
        docx_path=args.docx_path,
        anchor_text=anchor_text,
        items=items,
    )


def main():
    parser = argparse.ArgumentParser(description="HIT 学位论文 docx 自动化工具")
    sub = parser.add_subparsers(dest="command", required=True)

    # ── footer 子命令 ──────────────────────────────────────────────────────────
    p_footer = sub.add_parser("footer", help="设置 HIT 页码格式")
    p_footer.add_argument("docx_path", help="目标 .docx 文件路径")
    p_footer.add_argument(
        "--body-section",
        type=int,
        default=4,
        dest="body_section",
        help="正文（绪论）所在节编号，默认 4",
    )
    p_footer.set_defaults(func=cmd_footer)

    # ── insert-image 子命令 ────────────────────────────────────────────────────
    p_img = sub.add_parser("insert-image", help="批量在段落后插入图片")
    p_img.add_argument("docx_path", help="目标 .docx 文件路径")
    p_img.add_argument(
        "anchor_text",
        nargs="?",
        default=None,
        help="第一张图片的定位段落文字（使用 --from-json 时可省略）",
    )
    src = p_img.add_mutually_exclusive_group()
    src.add_argument(
        "--from-json",
        dest="from_json",
        metavar="JSON_FILE",
        help="从 JSON 文件读取 anchor_text 与 items",
    )
    src.add_argument(
        "--item",
        nargs="+",
        action="append",
        metavar=("PATH", "CAPTION"),
        help="一张图片：path caption [width height]，可重复",
    )
    p_img.set_defaults(func=cmd_insert_image)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
