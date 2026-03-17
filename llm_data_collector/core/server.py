from flask import Flask, request, jsonify
from typing import Dict, Any, List
import os
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
        """
        接收远程发送的 docx 文件 URL，下载并解析。
        请求体格式: {"url": "文件下载URL"}
        文件会被下载到临时目录，然后调用 parse_full_docx 进行解析。
        返回解析后的 JSON 数据。
        """
        try:
            data = request.get_json()
            if not data or 'url' not in data:
                return jsonify({"status": "error", "message": "Missing required field: url"}), 400

            file_url = data['url']

            # 下载文件到临时目录
            response = requests.get(file_url, timeout=30)
            response.raise_for_status()

            # 创建临时文件
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, "temp_docx_file.docx")

            with open(temp_path, 'wb') as f:
                f.write(response.content)

            # 解析 docx 文件
            parsed_result = parse_full_docx(temp_path)

            # 清理临时文件
            try:
                os.remove(temp_path)
            except Exception:
                pass

            return jsonify({
                "status": "success",
                "message": "File downloaded and parsed successfully",
                "data": parsed_result
            }), 200

        except requests.exceptions.RequestException as e:
            return jsonify({"status": "error", "message": f"Failed to download file: {str(e)}"}), 400
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 500

    @app.route('/health', methods=['GET'])
    def health():
        return jsonify({"status": "healthy"}), 200

    return app
