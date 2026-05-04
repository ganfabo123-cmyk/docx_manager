import requests

BASE_URL = "http://10.68.232.95:5001"


def get_docs(params):
    """
    输入:
        file_url (str): 原始 .doc/.docx 文件的下载地址

    输出:
        job_id        (str):  服务端任务 ID，后续所有节点都需要它
        download_url  (str):  无图版文档下载链接
        images        (list): 文档中的图片摘要列表，每项: {index, anchor_text, caption}
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
        raise RuntimeError(f"convert failed: {data.get(chr(39)+'message'+chr(39))}")

    images = data.get("images", [])

    return {
        "job_id":       data["job_id"],
        "download_url": BASE_URL + data["download_url"],
        "images":       images,
        "image_count":  len(images),
        "captions":     [img["caption"] for img in images],
    }


def insert_images(params):
    """
    输入:
        job_id         (str):  来自 get_docs
        single_images  (list): 单列排版图片，每项: {display_caption, original_caption, anchor_text}
        two_col_images (list): 双列排版图片，每项同上（必须为偶数张）
        chapter        (int):  单列图片章节编号，默认 1
        fig_start      (int):  单列图片起始编号，默认 1
        total_caption  (str, optional): 双列图片总图题，省略则服务端自动生成
        debug          (bool, optional): 双列排版 debug 截图开关，默认 False
        phases         (list, optional): 双列排版执行阶段，默认 [1,2,3,4,5]

    输出:
        single_download_url  (str): 单列图片插入后的文档下载链接（无单列图片时为空）
        two_col_download_url (str): 双列图片插入后的文档下载链接（无双列图片时为空）
        single_inserted      (int): 实际插入的单列图片数量
        two_col_inserted     (int): 实际插入的双列图片数量
    """
    job_id         = params["job_id"]
    single_items   = params.get("single_images", [])
    two_col_items  = params.get("two_col_images", [])

    result = {
        "single_download_url":  "",
        "two_col_download_url": "",
        "single_inserted":      0,
        "two_col_inserted":     0,
    }

    # 单列排版
    if single_items:
        resp = requests.post(
            f"{BASE_URL}/insert-image",
            json={
                "job_id":            job_id,
                "captions":          [img["display_caption"]  for img in single_items],
                "original_captions": [img["original_caption"] for img in single_items],
                "anchor_text":       single_items[0]["anchor_text"],
                "chapter":           params.get("chapter", 1),
                "fig_start":         params.get("fig_start", 1),
            },
            timeout=300,
            proxies={"http": None, "https": None},
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("status") != "ok":
            raise RuntimeError(f"insert-image failed: {data.get(chr(39)+'message'+chr(39))}")
        result["single_download_url"] = BASE_URL + data["download_url"]
        result["single_inserted"]     = len(single_items)

    # 双列排版
    if two_col_items:
        if len(two_col_items) % 2 != 0:
            raise ValueError(f"two-col layout requires an even number of images, got {len(two_col_items)}")
        resp = requests.post(
            f"{BASE_URL}/two-col",
            json={
                "job_id":            job_id,
                "captions":          [img["display_caption"]  for img in two_col_items],
                "original_captions": [img["original_caption"] for img in two_col_items],
                "anchor_text":       two_col_items[0]["anchor_text"],
                "total_caption":     params.get("total_caption") or None,
                "debug":             bool(params.get("debug", False)),
                "phases":            params.get("phases", [1, 2, 3, 4, 5]),
            },
            timeout=300,
            proxies={"http": None, "https": None},
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("status") != "ok":
            raise RuntimeError(f"two-col failed: {data.get(chr(39)+'message'+chr(39))}")
        result["two_col_download_url"] = BASE_URL + data["download_url"]
        result["two_col_inserted"]     = len(two_col_items)

    return result


def command_fixed_footer(params):
    """
    输入:
        job_id       (str): 来自 get_docs
        body_section (int): 正文（绪论）所在节编号，默认 4

    输出:
        download_url (str): 应用 HIT 页码格式后的最终文档下载链接
    """
    job_id       = params["job_id"]
    body_section = int(params.get("body_section", 4))

    resp = requests.post(
        f"{BASE_URL}/footer",
        json={
            "job_id":       job_id,
            "body_section": body_section,
        },
        timeout=120,
        proxies={"http": None, "https": None},
    )
    resp.raise_for_status()
    data = resp.json()

    if data.get("status") != "ok":
        raise RuntimeError(f"footer failed: {data.get(chr(39)+'message'+chr(39))}")

    return {
        "download_url": BASE_URL + data["download_url"],
    }
