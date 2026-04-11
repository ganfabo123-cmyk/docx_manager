import requests
from typing import Dict, Any, Optional, List

BASE_URL = "http://60.205.194.87:5001"


def send_request(endpoint: str, value: Any) -> Dict[str, Any]:
    try:
        response = requests.post(
            f"{BASE_URL}{endpoint}",
            json={"value": value},
            headers={"Content-Type": "application/json"}
        )
        response.raise_for_status()  # 确保请求成功
        result = response.json()
        return result
    except Exception as error:
        print(f"Error sending to {endpoint}: {error}")
        raise


def handler(params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    ret = {
        "success": True,
        "results": []
    }

    try:
        if params and "ref_citations" in params:
            ref_citations = params["ref_citations"]
            if isinstance(ref_citations, list) and len(ref_citations) > 0:
                citations = []
                for cit in ref_citations:
                    ref_id = cit.get("ref_id") or cit.get("refId") or cit.get("id")
                    if ref_id:
                        citations.append({
                            "ref_id": ref_id,
                            "before": cit.get("before", ""),
                            "after": cit.get("after", "")
                        })
                
                if citations:
                    result = send_request('/citations', citations)
                    ret["results"].append({"endpoint": "/citations", "status": result.get("status")})

    except Exception as error:
        ret["success"] = False
        ret["error"] = str(error)

    return ret


if __name__ == "__main__":
    # 示例调用
    test_params = {
        "ref_citations": [
            {"ref_id": 1, "before": "测试引用1", "after": ""},
            {"ref_id": 2, "before": "测试引用2", "after": ""}
        ]
    }
    
    result = handler(test_params)
    print(result)
