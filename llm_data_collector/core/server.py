from flask import Flask, request, jsonify
from typing import Dict, Any, List
import os
import sys
import traceback
import tempfile
import requests
from ..models.models import UserData, PageFooterConfig, TocEntry, Reference, Citation
from ..utils.parse_full_docx import parse_full_docx


class DataCollector:
    def __init__(self):
        self.user_data = UserData()

    def set_doc(self, doc: str):
        self.user_data._doc = doc

    def set_page_footer_config(self, configs: List[Dict[str, Any]]):
        self.user_data.page_footer_config = [
            PageFooterConfig(
                section=config["section"],
                style=config["style"],
                start=config["start"]
            )
            for config in configs
        ]

    def set_toc_mode(self, mode: str):
        self.user_data.toc_mode = mode

    def set_toc_entries(self, entries: List[Dict[str, Any]]):
        self.user_data.toc_entries = [
            TocEntry(
                title=entry["title"],
                level=entry["level"],
                page=entry["page"]
            )
            for entry in entries
        ]

    def add_content(self, content_item: Dict[str, Any]):
        self.user_data.content.append(content_item)

    def set_references(self, refs: List[Dict[str, Any]]):
        self.user_data.references = [
            Reference(
                id=ref["id"],
                text=ref["text"]
            )
            for ref in refs
        ]

    def set_citations(self, citations: List[Dict[str, Any]]):
        self.user_data.citations = [
            Citation(
                ref_id=cit["ref_id"],
                before=cit["before"],
                after=cit["after"]
            )
            for cit in citations
        ]

    def get_user_data(self) -> Dict[str, Any]:
        return self.user_data.to_dict()

    def reset(self):
        self.user_data = UserData()


collector = DataCollector()


def create_app():
    app = Flask(__name__)

    @app.route('/_doc', methods=['POST'])
    def receive_doc():
        try:
            data = request.get_json()
            if 'value' in data:
                collector.set_doc(data['value'])
            return jsonify({"status": "success", "message": "_doc received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/page_footer_config', methods=['POST'])
    def receive_page_footer_config():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], list):
                collector.set_page_footer_config(data['value'])
            return jsonify({"status": "success", "message": "page_footer_config received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/toc_mode', methods=['POST'])
    def receive_toc_mode():
        try:
            data = request.get_json()
            if 'value' in data:
                collector.set_toc_mode(data['value'])
            return jsonify({"status": "success", "message": "toc_mode received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/toc_entries', methods=['POST'])
    def receive_toc_entries():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], list):
                collector.set_toc_entries(data['value'])
            return jsonify({"status": "success", "message": "toc_entries received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/content_section', methods=['POST'])
    def receive_content_section():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                content_item = {"type": "section", **data['value']}
                collector.add_content(content_item)
            return jsonify({"status": "success", "message": "content_section received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/content_toc', methods=['POST'])
    def receive_content_toc():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                content_item = {"type": "toc", **data['value']}
                collector.add_content(content_item)
            return jsonify({"status": "success", "message": "content_toc received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/content_heading1', methods=['POST'])
    def receive_content_heading1():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                content_item = {"type": "heading1", **data['value']}
                collector.add_content(content_item)
            return jsonify({"status": "success", "message": "content_heading1 received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/content_heading2', methods=['POST'])
    def receive_content_heading2():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                content_item = {"type": "heading2", **data['value']}
                collector.add_content(content_item)
            return jsonify({"status": "success", "message": "content_heading2 received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/content_heading3', methods=['POST'])
    def receive_content_heading3():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                content_item = {"type": "heading3", **data['value']}
                collector.add_content(content_item)
            return jsonify({"status": "success", "message": "content_heading3 received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/content_body', methods=['POST'])
    def receive_content_body():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                content_item = {"type": "body", **data['value']}
                collector.add_content(content_item)
            return jsonify({"status": "success", "message": "content_body received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/content_table', methods=['POST'])
    def receive_content_table():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                content_item = {"type": "table", **data['value']}
                collector.add_content(content_item)
            return jsonify({"status": "success", "message": "content_table received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/content_formula', methods=['POST'])
    def receive_content_formula():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                content_item = {"type": "formula", **data['value']}
                collector.add_content(content_item)
            return jsonify({"status": "success", "message": "content_formula received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/content_image', methods=['POST'])
    def receive_content_image():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                content_item = {"type": "image", **data['value']}
                collector.add_content(content_item)
            return jsonify({"status": "success", "message": "content_image received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/references', methods=['POST'])
    def receive_references():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], list):
                collector.set_references(data['value'])
            return jsonify({"status": "success", "message": "references received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/citations', methods=['POST'])
    def receive_citations():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], list):
                collector.set_citations(data['value'])
            return jsonify({"status": "success", "message": "citations received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/get_data', methods=['GET'])
    def get_data():
        try:
            return jsonify({"status": "success", "data": collector.get_user_data()}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/reset', methods=['POST'])
    def reset():
        try:
            collector.reset()
            return jsonify({"status": "success", "message": "data reset"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/docx_send', methods=['POST'])
    def receive_docx():
        print("\n" + "="*30 + " 收到请求 " + "="*30)
        
        try:
            # 1. 解析请求数据
            data = request.get_json()
            print(f"[DEBUG 1/4] 请求 JSON 数据: {data}")

            if not data or 'url' not in data:
                print("[DEBUG ERROR] 缺少 url 字段")
                return jsonify({"status": "error", "message": "Missing required field: url"}), 400

            file_url = data['url']
            print(f"[DEBUG 2/4] 准备下载 URL: {file_url}")

            # 2. 下载文件
            try:
                # 加入 verify=False 如果是学校内部 HTTPS 证书有问题可以尝试，但先保持原样
                response = requests.get(file_url, timeout=30)
                response.raise_for_status()
                print(f"[DEBUG 2/4] 下载成功，文件大小: {len(response.content)} 字节")
            except requests.exceptions.RequestException as re:
                print(f"[DEBUG ERROR] 下载过程中出错: {str(re)}")
                return jsonify({"status": "error", "message": f"Download failed: {str(re)}"}), 400

            # 3. 写入临时文件
            temp_dir = tempfile.gettempdir()
            # 建议加个唯一标识，防止并发冲突
            temp_path = os.path.join(temp_dir, f"debug_{os.getpid()}.docx")
            print(f"[DEBUG 3/4] 正在写入临时文件: {temp_path}")

            with open(temp_path, 'wb') as f:
                f.write(response.content)
            print("[DEBUG 3/4] 临时文件写入完毕")

            # 4. 调用解析函数 (这里最容易报 500)
            print("[DEBUG 4/4] 开始进入 parse_full_docx 函数...")
            try:
                parsed_result = parse_full_docx(temp_path)
                print("[DEBUG 4/4] 解析函数执行成功")
            except Exception as parse_err:
                print("[DEBUG ERROR] parse_full_docx 内部崩溃了！")
                # 这一步非常重要，会让外层的 except 捕获到它
                raise parse_err

            # 5. 清理文件
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    print("[DEBUG] 临时文件已清理")
            except Exception as e:
                print(f"[DEBUG WARNING] 清理文件失败: {str(e)}")

            print("="*30 + " 处理完成 " + "="*30 + "\n")
            return jsonify({
                "status": "success",
                "message": "File downloaded and parsed successfully",
                "data": parsed_result
            }), 200

        except Exception as e:
            # 【核心调试代码】获取完整的报错堆栈
            error_type, error_value, error_trace = sys.exc_info()
            full_error = traceback.format_exc()
            
            print("\n" + "!"*20 + " 发生 500 错误 " + "!"*20)
            print(full_error)  # 在服务器黑窗口里打印完整错误
            print("!"*54 + "\n")

            # 返回给 HiAgent，让你在网页上就能看到错误详情
            return jsonify({
                "status": "error",
                "message": str(e),
                "error_type": str(error_type),
                "debug_trace": full_error  # 把这个字段发回给平台，一目了然
            }), 500

    @app.route('/health', methods=['GET'])
    def health():
        return jsonify({"status": "healthy"}), 200

    return app
