import requests

BASE_URL = "http://10.68.232.95:5001"


def handler(params):
    """
    输入:
        job_id       (str): 来自 get_docs 节点
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
        proxies={'http':None,'https':None},
        timeout=120,
    )
    resp.raise_for_status()
    data = resp.json()

    if data.get("status") != "ok":
        raise RuntimeError(f"footer failed: {data.get('message')}")

    return {
        "download_url": BASE_URL + data["download_url"],
    }
