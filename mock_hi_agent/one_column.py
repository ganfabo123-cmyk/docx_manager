import requests

BASE_URL = "http://10.68.232.95:5001"


def handler(params):
    """
    输入:
        job_id            (str):  来自 get_docs 节点
        configured_images (list): 来自 ask_user_for_image_config 节点的完整列表
                                  本节点只处理 layout == "single" 的项

    输出:
        download_url (str): 插入单列图片后的文档下载链接
        inserted     (int): 实际插入的图片数量
    """
    job_id            = params["job_id"]
    configured_images = params["configured_images"]

    single_items = [img for img in configured_images if img.get("layout") == "single"]

    if not single_items:
        return {
            "download_url": "",
            "inserted":     0,
        }

    resp = requests.post(
        f"{BASE_URL}/insert-image",
        json={
            "job_id":             job_id,
            "captions":           [img["display_caption"]   for img in single_items],
            "original_captions":  [img["original_caption"]  for img in single_items],
            "anchor_text":        single_items[0]["anchor_text"],
            "chapter":            params.get("chapter", 1),
            "fig_start":          params.get("fig_start", 1),
        },
        timeout=300,
        proxies={"http": None, "https": None},
    )
    resp.raise_for_status()
    data = resp.json()

    if data.get("status") != "ok":
        raise RuntimeError(f"insert-image failed: {data.get('message')}")

    return {
        "download_url": BASE_URL + data["download_url"],
        "inserted":     len(single_items),
    }

if __name__ == "__main__":
    result = handler({
        "job_id":    "5df9b96a1f924384b368af2296019e36",
        "chapter":   3,
        "fig_start": 1,
        "configured_images": [
            {
                "layout":           "single",
                "anchor_text":      "……",
                "original_caption": "(a) 气体静压轴承",
                "display_caption":  "(a) 气体静压轴承",
            },
            {
                "layout":           "single",
                "anchor_text":      "(a) 气体静压轴承",
                "original_caption": "(b) 气体动压轴承",
                "display_caption":  "(b) 气体动压轴承",
            },
            {
                "layout":           "single",
                "anchor_text":      "(b) 气体动压轴承",
                "original_caption": "(c) 气体动静压轴承",
                "display_caption":  "(c) 气体动静压轴承",
            },
            {
                "layout":           "single",
                "anchor_text":      "(c) 气体动静压轴承",
                "original_caption": "(d) 气体压膜轴承",
                "display_caption":  "(d) 气体压膜轴承",
            },
            {
                "layout":           "single",
                "anchor_text":      "（也可以按照下图范例书写）",
                "original_caption": "(a) 气体静压轴承",
                "display_caption":  "(a) 气体静压轴承",
            },
            {
                "layout":           "single",
                "anchor_text":      "(a) 气体静压轴承",
                "original_caption": "(b) 气体动压轴承",
                "display_caption":  "(b) 气体动压轴承",
            },
        ],
    })
    print(result)