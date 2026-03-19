from flask import Flask, request, jsonify
from typing import Dict, Any, List
import subprocess
import os
import sys
import traceback
import tempfile
import requests
from ..models.models import UserData, PageFooterConfig, TocEntry, Reference, Citation
from ..utils.parse_full_docx import parse_full_docx
import json
from pathlib import Path
class DataCollector:
    def __init__(self):
        self.user_data = UserData()
        self.toc_title = "目  录"
        self.image_defaults = {
            "width": 3.5,
            "align": "center",
            "ext": "png"
        }
        self.formula_defaults = {
            "label_prefix": "式"
        }

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

    def set_toc_title(self, title: str):
        self.toc_title = title

    def set_toc_entries(self, entries: List[Dict[str, Any]]):
        self.user_data.toc_entries = [
            TocEntry(
                title=entry["title"],
                level=entry["level"],
                page=entry["page"]
            )
            for entry in entries
        ]

    def set_image_defaults(self, defaults: Dict[str, Any]):
        self.image_defaults = defaults

    def set_formula_defaults(self, defaults: Dict[str, Any]):
        self.formula_defaults = defaults

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
        data = self.user_data.to_dict()
        data['toc_title'] = self.toc_title
        data['image_defaults'] = self.image_defaults
        data['formula_defaults'] = self.formula_defaults
        return data

    def reset(self):
        self.user_data = UserData()
        self.toc_title = "目  录"
        self.image_defaults = {
            "width": 3.5,
            "align": "center",
            "ext": "png"
        }
        self.formula_defaults = {
            "label_prefix": "式"
        }


collector = DataCollector()


def create_app(default_output_path=None): # 1. 允许传入默认输出路径
    app = Flask(__name__)
    
    # 页脚样式白名单
    VALID_FOOTER_STYLES = {
        'roman_lower_center', 'roman_upper_center',
        'arabic_center', 'arabic_dash',
        'arabic_page_x', 'arabic_slash', 'none'
    }

    @app.route('/save', methods=['POST'])
    def save_to_disk():
        try:
            # 尝试从 POST 请求体中获取文件名，如果没有则使用启动时的默认路径，再没有就存为默认名
            req_data = request.get_json(silent=True) or {}
            target_path = req_data.get('filename') or default_output_path or "collected_data.json"
            
            # 获取当前内存中的所有数据
            current_data = collector.get_user_data()
            
            # 确保目录存在
            output_path = Path(target_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 写入文件
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(current_data, f, ensure_ascii=False, indent=2)
                
            return jsonify({
                "status": "success", 
                "message": f"Data saved successfully to {target_path}",
                "path": str(output_path.absolute())
            }), 200
        except Exception as e:
            return jsonify({"status": "error", "message": f"Save failed: {str(e)}"}), 500


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
                # 验证页脚样式
                for config in data['value']:
                    if 'style' in config:
                        style = config['style']
                        if style not in VALID_FOOTER_STYLES:
                            return jsonify({"status": "error", "message": f"Invalid footer style: {style}. Valid styles are: {', '.join(VALID_FOOTER_STYLES)}"}), 400
                collector.set_page_footer_config(data['value'])
            return jsonify({"status": "success", "message": "page_footer_config received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/toc_title', methods=['POST'])
    def receive_toc_title():
        try:
            data = request.get_json()
            if 'value' in data:
                collector.set_toc_title(data['value'])
            return jsonify({"status": "success", "message": "toc_title received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/image_defaults', methods=['POST'])
    def receive_image_defaults():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                collector.set_image_defaults(data['value'])
            return jsonify({"status": "success", "message": "image_defaults received"}), 200
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/formula_defaults', methods=['POST'])
    def receive_formula_defaults():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], dict):
                collector.set_formula_defaults(data['value'])
            return jsonify({"status": "success", "message": "formula_defaults received"}), 200
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
        try:
            data = request.get_json()
            file_url = data.get('url', '')
            
            # 1. 下载文件
            response = requests.get(file_url, timeout=30)
            temp_dir = tempfile.gettempdir()
            # 初始下载的文件名（可能是 .doc）
            raw_path = os.path.join(temp_dir, f"input_{os.getpid()}.doc")
            
            with open(raw_path, 'wb') as f:
                f.write(response.content)
            print(f"[DEBUG] 文件已下载到: {raw_path}")

            # 2. 核心步骤：自动转换 .doc 为 .docx
            # 使用 libreoffice 进行转换
            try:
                print("[DEBUG] 正在尝试将 .doc 转换为 .docx...")
                # 命令解释：--headless 不启动界面，--convert-to 转换格式，--outdir 输出目录
                subprocess.run([
                    'libreoffice', '--headless', 
                    '--convert-to', 'docx', 
                    raw_path, 
                    '--outdir', temp_dir
                ], check=True, timeout=60)
                
                # 转换后的文件名会自动变成 .docx
                docx_path = raw_path.replace('.doc', '.docx')
                
                if not os.path.exists(docx_path):
                    raise Exception("LibreOffice 转换成功但未找到输出文件")
                    
                print(f"[DEBUG] 转换成功，新文件路径: {docx_path}")
            except Exception as e:
                print(f"[DEBUG ERROR] 转换失败: {str(e)}")
                return jsonify({"status": "error", "message": f"Conversion failed: {str(e)}"}), 500

            # 3. 解析转换后的 .docx
            parsed_result = parse_full_docx(docx_path)

            # 4. 清理所有临时文件
            for p in [raw_path, docx_path]:
                if os.path.exists(p):
                        os.remove(p)
            # [DEBUG] 打印转换和预览 (你已经看到了，说明到这里都没问题)
            result_str = str(parsed_result)
            print(f"[DEBUG] 准备发送的数据总长度: {len(result_str)} 字符")

            # --- 暴力修改开始 ---
            try:
                # 使用 json.dumps 手动转成字符串，确保没有编码问题
                # ensure_ascii=False 保证中文不乱码
                json_body = json.dumps({
                    "status": "success",
                    "message": "File downloaded and parsed successfully",
                    "data": parsed_result
                }, ensure_ascii=False)
                
                print(f"[DEBUG] 最终 JSON 字节长度: {len(json_body.encode('utf-8'))}")

                # 直接返回 Flask Response 对象，不经过 jsonify 
                from flask import Response
                return Response(json_body, content_type='application/json; charset=utf-8'), 200

            except Exception as json_err:
                print(f"[DEBUG ERROR] JSON 序列化失败: {str(json_err)}")
                return jsonify({"status": "error", "message": f"Serialization error: {str(json_err)}"}), 500
        except Exception as e:
            # 保留你之前的 traceback 调试代码...
            return jsonify({"status": "error", "message": str(e)}), 500
    @app.route('/health', methods=['GET'])
    def health():
        return jsonify({"status": "healthy"}), 200

    return app
