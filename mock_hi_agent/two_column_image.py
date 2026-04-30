import requests

BASE_URL = "http://10.68.232.95:5001"


def handler(params):
    """
    输入:
        job_id            (str):  来自 get_docs 节点
        configured_images (list): 来自 ask_user_for_image_config 节点的完整列表
                                  本节点只处理 layout == "two-col" 的项（必须恰好 2 张）
        total_caption     (str, optional): 总图题，省略则服务端自动生成
        debug             (bool, optional): 是否开启 debug 截图，默认 False
        phases            (list, optional): 要执行的阶段，默认 [1,2,3,4,5]

    输出:
        download_url (str): 插入双列图片后的文档下载链接
        inserted     (int): 实际插入的图片数量（0 或 2）
    """
    job_id            = params["job_id"]
    configured_images = params["configured_images"]
    total_caption     = params.get("total_caption") or None
    debug             = bool(params.get("debug", False))
    phases            = params.get("phases", [1, 2, 3, 4, 5])

    two_col_items = [img for img in configured_images if img.get("layout") == "two-col"]

    if not two_col_items:
        return {
            "download_url": "",
            "inserted":     0,
        }

    if len(two_col_items)%2 != 0:
        raise ValueError(
            f"two-col layout requires exactly double images, got {len(two_col_items)}"
        )

    resp = requests.post(
        f"{BASE_URL}/two-col",
        json={
            "job_id":            job_id,
            "captions":          [img["display_caption"]  for img in two_col_items],
            "original_captions": [img["original_caption"] for img in two_col_items],
            "anchor_text":       two_col_items[0]["anchor_text"],
            "total_caption":     total_caption,
            "debug":             debug,
            "phases":            phases,
        },
        timeout=300,
        proxies={"http": None, "https": None},
    )
    resp.raise_for_status()
    data = resp.json()

    if data.get("status") != "ok":
        raise RuntimeError(f"two-col failed: {data.get('message')}")

    return {
        "download_url": BASE_URL + data["download_url"],
        "inserted":     2,
    }


if __name__ == "__main__":
    result = handler({
        "job_id": "7c88ce6423864f74aa948ba831f00d8e",
        "total_caption": None,   # None → 服务端自动生成 "图A和图B"
        "debug": False,
        "phases": [1, 2, 3, 4, 5],
        "configured_images": [
            {
                "layout":           "two-col",
                "anchor_text":      "……",
                "original_caption": "(a) 气体静压轴承",
                "display_caption":  "(a) 气体静压轴承",
            },
            {
                "layout":           "two-col",
                "anchor_text":      "(a) 气体静压轴承",
                "original_caption": "(b) 气体动压轴承",
                "display_caption":  "(b) 气体动压轴承",
            },
        ],
    })
    print(result)
