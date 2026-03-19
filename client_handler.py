import requests
import json
import traceback

def handler(params):
    file_url = params.get('url', '')
    target_api = "http://60.205.194.87:5001/parse_docx_partly"
    
    # 预定义所有输出变量名
    final_output = {
        "res_citations":[],
        "status": "initializing"
    }
    
    try:
        # 发送请求
        response = requests.post(target_api, json={"url": file_url}, timeout=120)
        
        if response.status_code != 200:
            final_output["status"] = "error"
            final_output["message"] = f"Server error: {response.status_code}"
            return final_output
            
        full_result = response.json()
        
        # 检查服务器返回状态
        if full_result.get("status") != "success":
            final_output["status"] = "error"
            final_output["message"] = full_result.get("message", "Unknown error")
            return final_output
        
        # 获取 data 对象（包含 references 和 body）
        data = full_result.get("data", {})
        citations:dict = full_result.get("data",{}).get("citations")
        # 处理 references 数组
        references_list = data.get("references", [])

        citations_with_references = citations.copy()
        for item in citations_with_references:
            for reference in references_list:
                rid = item.get("ref_id","")
                if f"［{rid}］" in reference.get("value",""):
                    item['reference'] = reference.get("value","")

        final_output["res_citations"] = citations_with_references 
        final_output["status"] = "success"
        return final_output

    except Exception as e:
        # 打印详细错误到平台日志以便排查
        print(f"Code Node Error Trace:\n{traceback.format_exc()}")
        final_output["status"] = "error"
        final_output["message"] = str(e)
        # 即使报错也要返回所有预设的 Key，防止平台二次报错
        return final_output

if __name__ == "__main__":
    handler({'url':'https://agent.hit.edu.cn/api/proxy/down?Action=Download&Version=2022-01-01&Path=upload%2Ffull%2F66%2F9b%2Fc70cedf0e0971155db2a4849dd4b4198757825bcb8280ff32e71b4cc64ff&IsAnonymous=true'})