import requests

BASE_URL = "http://10.68.232.95:5001"


def handler(params):
    """
    输入:
        file_url (str): 原始 .doc/.docx 文件的下载地址

    输出:
        job_id        (str):  服务端任务 ID，后续所有路由都需要它
        download_url  (str):  无图版文档下载链接（无图片、无页码）
        images        (list): 文档中的图片摘要列表
                              每项: {index, anchor_text, caption}
        image_count   (int):  图片总数
    """
    file_url = params["file_url"]

    resp = requests.post(
        "http://localhost:5001/convert",
        json={"url": file_url},
        
        proxies={"http": None, "https": None},
    )
    resp.raise_for_status()
    data = resp.json()

    if data.get("status") != "ok":
        raise RuntimeError(f"convert failed: {data.get('message')}")

    images = data.get("images", [])

    return {
        "job_id":       data["job_id"],
        "download_url": BASE_URL + data["download_url"],
        "images":       images,
        "image_count":  len(images),
        "captions":     [img["caption"] for img in images],
    }
