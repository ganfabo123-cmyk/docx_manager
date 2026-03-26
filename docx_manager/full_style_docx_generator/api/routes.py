from flask import request, jsonify
import traceback
import os
import tempfile
import json
import requests
from core.document_processor import parse_document_content, identify_short_blocks, generate_docx_document
from core.file_parser import parse_file

try:
    import pythoncom
    import win32com.client as win32
    WORD_AVAILABLE = True
except ImportError:
    print("win32com not installed, .doc conversion will be disabled")
    WORD_AVAILABLE = False

def register_routes(app):
    @app.route('/parse-document', methods=['POST'])
    def parse_document():
        try:
            data = request.json
            document_content = data.get('content', '')
            
            blocks = parse_document_content(document_content)
            
            return jsonify({'blocks': blocks})
        except Exception as e:
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/parse-txt-file', methods=['POST'])
    def parse_txt_file():
        temp_path = None
        try:
            data = request.json
            url = data.get('url', '')
            
            if not url:
                return jsonify({'error': 'No URL provided'}), 400
            
            print(f"Downloading txt file from: {url}")
            
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"input_{os.getpid()}.txt")
            
            with open(temp_path, 'wb') as f:
                f.write(response.content)
            print(f"Downloaded file to: {temp_path}")
            
            blocks = parse_file(temp_path)
            print(f"Parsed {len(blocks)} blocks from file")
            
            if not blocks:
                return jsonify({'error': 'Failed to parse file'}), 500
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            output_path = os.path.join(data_dir, 'parsed_blocks.json')
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(blocks, f, ensure_ascii=False, indent=2)
            
            print(f"Saved parsed blocks to: {output_path}")
            
            return jsonify({'status': 'success', 'message': 'File parsed and blocks saved'}), 200
        except Exception as e:
            print(f"Error parsing txt file: {e}")
            print(traceback.format_exc())
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500
        finally:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
                print(f"Cleaned up temporary file: {temp_path}")

    def convert_doc_to_docx(doc_path: str) -> str:
        """
        Convert .doc file to .docx using win32com
        """
        temp_dir = tempfile.gettempdir()
        base_name = os.path.splitext(os.path.basename(doc_path))[0]
        docx_path = os.path.join(temp_dir, f"{base_name}.docx")
        
        abs_doc_path = os.path.abspath(doc_path)
        abs_docx_path = os.path.abspath(docx_path)
        
        print(f"[DEBUG] Converting .doc to .docx: {abs_doc_path} -> {abs_docx_path}")
        
        pythoncom.CoInitialize()
        word = None
        doc = None
        
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = False
            
            doc = word.Documents.Open(abs_doc_path)
            doc.SaveAs2(abs_docx_path, FileFormat=16)  # 16 = wdFormatXMLDocument
            
            print(f"[DEBUG] Conversion successful: {abs_docx_path}")
            
            if not os.path.exists(abs_docx_path):
                raise Exception("Word conversion completed but docx file not found")
            
            return abs_docx_path
            
        finally:
            if doc:
                doc.Close(False)
            if word:
                word.Quit()
            pythoncom.CoUninitialize()

    @app.route('/parse-docx-file', methods=['POST'])
    def parse_docx_file():
        temp_path = None
        converted_path = None
        try:
            data = request.json
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
                print(f"[DEBUG] 正在 Windows 环境下调用原生 Word 转换: {raw_path}")
                
                # 1. 准备路径（Word 必须使用绝对路径，且 Windows 路径需标准化）
                abs_raw_path = os.path.abspath(raw_path)
                base_name = os.path.splitext(os.path.basename(abs_raw_path))[0]
                docx_path = os.path.join(temp_dir, f"{base_name}.docx")
                abs_docx_path = os.path.abspath(docx_path)

                # 2. 初始化 COM 接口（如果在 Flask/FastAPI 等异步框架中运行，这是必须的）
                pythoncom.CoInitialize()

                word = None
                doc = None
                try:
                    # 3. 启动 Word 进程
                    # EnsureDispatch 会自动生成缓存，效率更高
                    word = win32.gencache.EnsureDispatch('Word.Application')
                    word.Visible = False  # 不显示 Word 界面
                    word.DisplayAlerts = False  # 不弹窗确认

                    # 4. 打开文档
                    doc = word.Documents.Open(abs_raw_path)

                    # 5. 另存为 docx
                    # FileFormat=16 代表 wdFormatXMLDocument (即 docx)
                    doc.SaveAs2(abs_docx_path, FileFormat=16)
                    
                    print(f"[DEBUG] 原生 Word 转换成功: {abs_docx_path}")

                finally:
                    # 6. 无论成功失败，必须关闭文档并退出，否则后台会堆积一大堆 winword.exe
                    if doc:
                        doc.Close(False)
                    # 如果你并发量大，建议把 word 对象写成全局单例，不要每次都 Quit
                    # 但如果是脚本运行，建议 Quit
                    if word:
                        word.Quit()
                    # 释放 COM 资源
                    pythoncom.CoUninitialize()

                if not os.path.exists(abs_docx_path):
                    raise Exception("Word 运行结束但未生成 docx 文件")

            except Exception as e:
                print(f"[DEBUG ERROR] Windows 原生转换失败: {str(e)}")
                # 这里可以加一个兜底，如果 Word 挂了，尝试直接改名（如果是伪 doc 的话）
                return jsonify({"status": "error", "message": f"Windows Word conversion failed: {str(e)}"}), 500

            blocks = parse_file(docx_path)
            
            print(f"Parsed {len(blocks)} blocks from file")
            
            if not blocks:
                return jsonify({'error': 'Failed to parse file'}), 500
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            output_path = os.path.join(data_dir, 'parsed_blocks.json')
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(blocks, f, ensure_ascii=False, indent=2)
            
            print(f"Saved parsed blocks to: {output_path}")
            
            return jsonify({'status': 'success', 'message': 'File parsed and blocks saved'}), 200
        except Exception as e:
            print(f"Error parsing docx file: {e}")
            print(traceback.format_exc())
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500
        finally:
            # Clean up temporary files
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
                print(f"Cleaned up temporary file: {temp_path}")
            if converted_path and os.path.exists(converted_path):
                os.remove(converted_path)
                print(f"Cleaned up converted file: {converted_path}")

    @app.route('/identify-short-blocks', methods=['GET'])
    def identify_short_blocks_endpoint():
        try:
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            blocks_path = os.path.join(data_dir, 'parsed_blocks.json')
            
            if not os.path.exists(blocks_path):
                return jsonify({'error': 'No parsed blocks found'}), 404
            
            with open(blocks_path, 'r', encoding='utf-8') as f:
                blocks = json.load(f)
            
            short_blocks = identify_short_blocks(blocks)
            
            print(f"Identified {len(short_blocks)} short blocks from {len(blocks)} total blocks")
            
            return jsonify({'short_blocks': short_blocks})
        except Exception as e:
            print(f"Error identifying short blocks: {e}")
            print(traceback.format_exc())
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/analyze-styles', methods=['POST'])
    def analyze_styles():
        try:
            data = request.json
            styles = data.get('styles', [])
            
            if not styles:
                return jsonify({'error': 'No styles provided'}), 400
            
            print(f"Received {len(styles)} styles from model")
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            output_path = os.path.join(data_dir, 'parsed_styles.json')
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(styles, f, ensure_ascii=False, indent=2)
            
            print(f"Saved parsed styles to: {output_path}")
            
            return jsonify({'status': 'success', 'message': 'Styles saved successfully'}), 200
        except Exception as e:
            print(f"Error saving styles: {e}")
            print(traceback.format_exc())
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/generate-document', methods=['POST'])
    def generate_document():
        try:
            data = request.json
            parsed_styles = data.get('parsed_styles', [])
            
            if not parsed_styles:
                return jsonify({'error': 'No parsed styles provided'}), 400
            
            downloads_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'download')
            os.makedirs(downloads_dir, exist_ok=True)
            
            output_filename = 'generated_document.docx'
            output_path = os.path.join(downloads_dir, output_filename)
            
            success = generate_docx_document(parsed_styles, output_path)
            
            if not success:
                return jsonify({'error': 'Failed to generate docx document'}), 500
            
            print(f"Generated docx document saved to: {output_path}")
            
            return jsonify({'status': 'success', 'message': 'Document generated successfully', 'file_path': output_path}), 200
        except Exception as e:
            print(f"Error generating document: {e}")
            print(traceback.format_exc())
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500