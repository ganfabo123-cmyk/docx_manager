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
        
        if params and "filename" in params:
            request_body["filename"] = params["filename"]

        response = requests.post(
            f"{BASE_URL}/save",
            json=request_body,
            headers={"Content-Type": "application/json"}
        )
        
        result = response.json()
        ret["result"] = result
        
        if result.get("status") != "success":
            ret["success"] = False
            
    except Exception as error:
        ret["success"] = False
        ret["error"] = str(error)

    return ret

if __name__ == "__main__":
    handler()