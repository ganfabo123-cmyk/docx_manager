from flask import Flask, request, jsonify, send_file
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

TEMPLATE_DOCX_PATH = "data\\full_template_v6.docx"
OUTPUT_DOCX_PATH = "download\\formatted_output.docx"

class DataCollector:
    def __init__(self):
        self.user_data = UserData()
        # 初始化默认页脚配置
        self.user_data.page_footer_config = [
            PageFooterConfig(
                section="frontmatter",
                style="roman_lower_center",
                start=1
            ),
            PageFooterConfig(
                section="mainmatter",
                style="arabic_dash",
                start=1
            )
        ]
        self.toc_title = "目  录"
        self.image_defaults = {
            "width": 3.5,
            "align": "center",
            "ext": "png"
        }
        self.formula_defaults = {
            "label_prefix": "式"
        }
        self.docx_infos = []

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

    def get_full_config(self) -> Dict[str, Any]:
        """获取完整的配置对象"""
        from pathlib import Path
        import json
        
        # 读取现有配置文件
        config_path = Path(__file__).parent.parent / 'utils' / 'user_config.json'
        base_config = {}
        if config_path.exists():
            with open(config_path, 'r', encoding='utf-8') as f:
                base_config = json.load(f)
        #print(base_config)
        # 合并用户设置的配置
        config = {
            "_doc": base_config.get("_doc", "用户配置文件 - 用于补全转换过程中缺失的信息"),
            "_tips": base_config.get("_tips", {
                "page_footer_config": "页脚配置，section 可选: frontmatter, mainmatter",
                "style_options": "frontmatter: roman_lower_center(小写罗马居中), roman_upper_right(大写罗马右对齐); mainmatter: arabic_center(阿拉伯数字居中), arabic_dash(阿拉伯数字带横线)",
                "toc_mode": "manual(手动目录) 或 auto(自动目录)",
                "image_defaults": "图片默认配置，当源数据缺失时使用",
                "formula_defaults": "公式默认配置"
            }),
            "page_footer_config": [
                {
                    "section": item.section,
                    "style": item.style,
                    "start": item.start
                }
                for item in self.user_data.page_footer_config
            ],
            "toc_mode": base_config.get("toc_mode", "manual"),
            "toc_title": self.toc_title,
            "toc_title_exclude": base_config.get("toc_title_exclude", True),
            "image_defaults": self.image_defaults,
            "formula_defaults": self.formula_defaults,
            "section_toc_exclude": base_config.get("section_toc_exclude", {
                "abstract": True,
                "abstract_en": True,
                "conclusion": True,
                "acknowledgement": True,
                "references": True
            }),
            "heading_toc_exclude_default": base_config.get("heading_toc_exclude_default", False),
            "citations": [
                {
                    "ref_id": cit.ref_id,
                    "before": cit.before,
                    "after": cit.after
                }
                for cit in self.user_data.citations
            ] if self.user_data.citations else []
        }
        
        return config

    def save_config(self) -> bool:
        """保存配置到文件"""
        from pathlib import Path
        import json
        
        config_path = Path(__file__).parent.parent / 'utils' / 'user_config.json'
        config = self.get_full_config()
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False

    def reset(self):
        self.user_data = UserData()
        # 重置默认页脚配置
        self.user_data.page_footer_config = [
            PageFooterConfig(
                section="frontmatter",
                style="roman_lower_center",
                start=1
            ),
            PageFooterConfig(
                section="mainmatter",
                style="arabic_dash",
                start=1
            )
        ]
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
    
    @app.route('/save', methods=['POST'])
    def save_to_disk():
        try:
            # 尝试从 POST 请求体中获取文件名，如果没有则使用启动时的默认路径，再没有就存为默认名
            req_data = request.get_json(silent=True) or {}
            
            # 默认为 user_config.json 路径
            default_path = Path(__file__).parent.parent / 'utils' / 'user_config.json'
            target_path = req_data.get('filename') or default_output_path or str(default_path)
            
            # 获取当前内存中的所有数据
            current_data = collector.get_full_config()
            
            # 确保目录存在
            output_path = Path(target_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            #save之前在控制台把current data数据打印出来
            print(f"[DEBUG] Current data to save: {json.dumps(current_data, ensure_ascii=False, indent=2)}")
            # 写入文件
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(current_data, f, ensure_ascii=False, indent=2)
            print(f'Data saved successfully to {output_path}')
            return jsonify({
                "status": "success", 
                "message": f"Data saved successfully to {output_path}",
                "path": str(output_path.absolute())
            }), 200
        except Exception as e:
            print(f"[DEBUG] Current data to save: {current_data}")
            print(f"[ERROR] Save failed: {str(e)}")
            print(traceback.format_exc())
            return jsonify({"status": "error", "message": f"Save failed: {str(e)}"}), 500

    @app.route('/citations', methods=['POST'])
    def receive_citations():
        try:
            data = request.get_json()
            if 'value' in data and isinstance(data['value'], list):
                collector.set_citations(data['value'])
            return jsonify({"status": "success", "message": "citations received"}), 200
        except Exception as e:
            print(f"[ERROR] Citations receive failed: {str(e)}")
            print(traceback.format_exc())
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/recieve_right_style_docx', methods=['POST'])
    def receive_need_docx():
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


            # 处理解析结果：
            # 1. 找到 '参考文献' heading1 之后的 body 元素作为 references
            # 2. 其他 body 元素作为 body
            references = []
            bodies = []
            found_references = False
            collector.docx_infos = parsed_result.get("docx_infos")
            for item in parsed_result.get("docx_infos"):
                item_type = item.get('type', '')
                item_value = item.get('value', '')
                
                # 检测是否遇到'参考文献'标题
                if item_type == 'heading1' and item_value == '参考文献':
                    found_references = True
                    continue
                
                # 收集元素
                if item_type == 'body':
                    if found_references:
                        # 参考文献后的 body 放入 references
                        references.append(item)
                    else:
                        # 其他 body 放入 bodies
                        bodies.append(item)
            
            # 构建返回数据
            result_data = {
                'references': references,
                'body': bodies,
                'citations': parsed_result.get("citations")
            }

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
                    "data": result_data
                }, ensure_ascii=False)
                
                print(f"[DEBUG] 最终 JSON 字节长度: {len(json_body.encode('utf-8'))}")
                print(f"[DEBUG] 提取到 {len(references)} 个 references, {len(bodies)} 个 body")

                # 直接返回 Flask Response 对象，不经过 jsonify 
                from flask import Response
                return Response(json_body, content_type='application/json; charset=utf-8'), 200

            except Exception as json_err:
                print(f"[DEBUG ERROR] JSON 序列化失败: {str(json_err)}")
                print(traceback.format_exc())
                return jsonify({"status": "error", "message": f"Serialization error: {str(json_err)}"}), 500
        except Exception as e:
            # 保留你之前的 traceback 调试代码...
            print(f"[ERROR] Recieve right style docx failed: {str(e)}")
            print(traceback.format_exc())
            return jsonify({"status": "error", "message": str(e)}), 500

    @app.route('/generate_user_data', methods=['POST'])
    def generate_user_data():
        try:
            # 保存配置到文件
            if not collector.save_config():
                return jsonify({"status": "error", "message": "Failed to save config"}), 500
            
            # 读取 DOCX 路径
            data = request.get_json()

            
            # 调用 generate_user_data 函数
            from llm_data_collector.utils.generate_user_data import generate_user_data_from_file, save_user_data
            
            result = generate_user_data_from_file(collector.docx_infos)
            
            # 保存用户数据 JSON
            output_path = str(data / 'generated_user_data.json')
            save_user_data(result, output_path)
            
            # 调用 process 函数生成格式化的 DOCX
            from docx_fixer.api import process
            
            # 确定模板路径和数据路径
            template_path = TEMPLATE_DOCX_PATH  # 使用原始 DOCX 作为模板
            data_path = output_path   # 使用生成的 JSON 作为数据
            formatted_output_path = OUTPUT_DOCX_PATH
            
            # 调用 process 函数（api_key 设为 None）
            process(template_path, data_path, formatted_output_path, None)
            
            # 返回下载链接
            filename = os.path.basename(formatted_output_path)
            download_url = f"/download/{filename}"
            
            return jsonify({
                "status": "success", 
                "message": "Formatted DOCX generated successfully",
                "download_url": download_url,
                "output_path": formatted_output_path,
                "user_data_path": output_path
            }), 200
        except Exception as e:
            print(f"[ERROR] Generate user data failed: {str(e)}")
            print(traceback.format_exc())
            return jsonify({"status": "error", "message": str(e)}), 400

    @app.route('/download/<filename>')
    def download_file(filename):
        try:
            from pathlib import Path
            
            # 从当前工作目录的 data 文件夹中查找文件
            file_path = Path('data') / filename
            
            if not file_path.exists():
                # 如果不在 data 文件夹，尝试在 docx_path 的父目录中查找
                file_path = Path.cwd() / filename
            
            if not file_path.exists():
                return jsonify({"status": "error", "message": "File not found"}), 404
            
            return send_file(str(file_path), as_attachment=True, download_name=filename)
        except Exception as e:
            print(f"[ERROR] Download file failed: {str(e)}")
            print(traceback.format_exc())
            return jsonify({"status": "error", "message": str(e)}), 500

    @app.route('/health', methods=['GET'])
    def health():
        return jsonify({"status": "healthy"}), 200

    return app
