import requests
from typing import Dict, Any, Optional

BASE_URL = "http://10.68.186.62:5001"


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
            
    except Exception as error:
        ret["success"] = False
        ret["error"] = str(error)

    return ret
