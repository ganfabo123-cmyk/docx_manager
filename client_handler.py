import requests
from typing import Dict, Any, Optional

BASE_URL = "http://localhost:5001"


def handler(params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    ret = {
        "success": True,
        "result": None
    }

    try:
        request_body = {}
        
        if params:
            if "docx_path" in params:
                request_body["docx_path"] = params["docx_path"]
            if "output_path" in params:
                request_body["output_path"] = params["output_path"]

        response = requests.post(
            f"{BASE_URL}/generate_user_data",
            json=request_body,
            headers={"Content-Type": "application/json"}
        )
        
        result = response.json()
        ret["result"] = result
        
        if result.get("status") != "success":
            ret["success"] = False
        else:
            # 如果成功，添加完整下载链接
            if "download_url" in result:
                ret["download_url"] = f"{BASE_URL}{result['download_url']}"
            if "output_path" in result:
                ret["output_path"] = result["output_path"]
            if "user_data_path" in result:
                ret["user_data_path"] = result["user_data_path"]
            
    except Exception as error:
        ret["success"] = False
        ret["error"] = str(error)

    return ret


if __name__ == "__main__":
    result = handler()
    print(result)