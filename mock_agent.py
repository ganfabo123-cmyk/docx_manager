import requests
import json
import traceback

def handler(params):
    file_url = params.get('url', '')
    target_api = "http://60.205.194.87:5001/docx_send"
    
    # 1. 预定义所有你在 UI 界面中创建的输出变量名
    # 即使文档里没有这些类型，我们也必须返回空列表 [] 给平台
    final_output = {
        "res_all_ids": [],
        "res_body": [],
        "res_heading1": [],
        "res_heading2": [],
        "res_heading3": [],
        "res_toc1": [],
        "res_toc2": [],
        "res_toc3": [],
        "res_reference": [],
        "res_image": [],
        "res_formula": [],
        "res_table": [],
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
        data_list = full_result.get("data", [])
        
        # 2. 填充数据
        for item in data_list:
            t = item.get('type')
            cid = item.get('chunk_id')
            if not t or not cid: continue
            
            # 记录全局 ID
            final_output["res_all_ids"].append(str(cid))
            
            # 提取内容并强制转为字符串（防止校验失败）
            val = ""
            if t == 'table':
                # 表格转为 JSON 字符串
                val = json.dumps(item.get('data', []), ensure_ascii=False)
            elif 'value' in item:
                val = str(item['value'])
            elif 'text' in item:
                val = str(item['text'])
            
            # 构造变量名
            var_name = f"res_{t}"
            
            # 只有当这个类型在我们预定义的 final_output 中时才进行 append
            if var_name in final_output:
                final_output[var_name].append({
                    "cid": str(cid),
                    "val": val
                })
        
        final_output["status"] = "success"
        return final_output

    except Exception as e:
        # 打印详细错误到平台日志以便排查
        print(f"Code Node Error Trace:\n{traceback.format_exc()}")
        final_output["status"] = "error"
        # 即使报错也要返回所有预设的 Key，防止平台二次报错
        return final_output

if __name__ == "__main__":
    handler({"name":"docx","url":"https://agent.hit.edu.cn/api/proxy/down?Action=Download&Version=2022-01-01&Path=upload%2Ffull%2F66%2F9b%2Fc70cedf0e0971155db2a4849dd4b4198757825bcb8280ff32e71b4cc64ff&IsAnonymous=true"})