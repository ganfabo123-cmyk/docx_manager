from flask import request, jsonify
from flask import request, jsonify, send_from_directory
import time  # 顺便引入 time，用来生成防重名的文件名
import traceback
import os
import tempfile
import json
import requests
from core.document_processor import identify_short_blocks, generate_docx_document
from core.file_parser import (
    parse_file,
    parse_docx_to_json,
    extract_text_from_parsed_json,
    backfill_styles_to_json,
    restore_docx_from_json,
    parse_text_to_json,
    remove_markdown_symbols
)

try:
    import pythoncom
    import win32com.client as win32
    WORD_AVAILABLE = True
except ImportError:
    print("win32com not installed, .doc conversion will be disabled")
    WORD_AVAILABLE = False

def register_routes(app):

    @app.errorhandler(Exception)
    def handle_exception(e):
        print(f"❌ [全局异常] {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return {"error": "Internal Server Error"}, 500

    @app.route('/parse-txt-file', methods=['POST'])
    def parse_txt_file():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /parse-txt-file")
        temp_path = None
        try:
            data = request.json or {}
            print(f"📦 [PAYLOAD] 原始入参: {json.dumps(data, ensure_ascii=False)[:300]} {'...' if len(str(data))>300 else ''}")
            
            # 兼容 {"url": "..."} 和 {"txt": {"url": "..."}} 两种格式
            url = data.get('url', '')
            if not url and isinstance(data.get('txt'), dict):
                url = data.get('txt').get('url', '')
                
            remove_md = data.get('remove_markdown', True)
            
            print(f"🔍 [PARSE] 提取到的 URL: {url}")
            print(f"⚙️  [PARSE] 是否移除 Markdown: {remove_md}")
            
            if not url:
                print("❌ [ERROR] 缺少 URL 参数")
                return jsonify({'error': 'No URL provided'}), 400
            
            print(f"⬇️  [ACTION] 开始下载 txt 文件...")
            
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"input_{os.getpid()}.txt")
            
            with open(temp_path, 'wb') as f:
                f.write(response.content)
            print(f"✅ [SUCCESS] 文件下载成功，保存至: {temp_path} (大小: {len(response.content)} bytes)")
            
            print(f"🔄 [ACTION] 开始解析 TXT 文件...")
            result = parse_file(temp_path, remove_md=remove_md)
            
            if not result:
                print("❌ [ERROR] 文件解析失败，返回空结果")
                return jsonify({'error': 'Failed to parse file'}), 500
            
            blocks = result.get("text_elements", [])
            print(f"📊 [INFO] 解析完成，共提取 {len(blocks)} 个文本块")
            
            #short_elements = {"text_elements":[e for e in blocks if len(e.get('content','')) < 30]}
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            output_path = os.path.join(data_dir, 'parsed_blocks.json')
            full_json_path = os.path.join(data_dir, 'full_parsed.json')

            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)

            with open(full_json_path, 'w', encoding='utf-8') as f:
                json.dump(blocks, f, ensure_ascii=False, indent=2)
            
            print(f"💾 [SAVE] 解析结果已保存至: {output_path}")
            print("🏁 [API OUT] 请求处理成功返回 200")
            
            return jsonify({'status': 'success', 'message': 'File parsed and blocks saved', 'text_count': len(blocks)}), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /parse-txt-file 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500
        finally:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
                print(f"🧹 [CLEANUP] 清理临时文件: {temp_path}")

    def convert_doc_to_docx(doc_path: str) -> str:
        # 内部函数保持原样，保留了你原来的 DEBUG print
        temp_dir = tempfile.gettempdir()
        base_name = os.path.splitext(os.path.basename(doc_path))[0]
        docx_path = os.path.join(temp_dir, f"{base_name}.docx")
        
        abs_doc_path = os.path.abspath(doc_path)
        abs_docx_path = os.path.abspath(docx_path)
        
        print(f"🛠️ [DEBUG] 正在将 .doc 转换为 .docx: {abs_doc_path} -> {abs_docx_path}")
        
        pythoncom.CoInitialize()
        word = None
        doc = None
        
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = False
            
            doc = word.Documents.Open(abs_doc_path)
            doc.SaveAs2(abs_docx_path, FileFormat=16)
            
            print(f"✅ [DEBUG] 转换成功: {abs_docx_path}")
            
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
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /parse-docx-file")
        temp_path = None
        converted_path = None
        try:
            data = request.json or {}
            print(f"📦 [PAYLOAD] 原始入参: {json.dumps(data, ensure_ascii=False)[:300]} {'...' if len(str(data))>300 else ''}")
            
            # 兼容 {"url": "..."} 和 {"docx": {"url": "..."}} 两种格式
            file_url = data.get('url', '')
            remove_md = data.get('remove_markdown', True)
            if not file_url and isinstance(data.get('docx'), dict):
                file_url = data.get('docx').get('url', '')
                
            print(f"🔍 [PARSE] 提取到的 URL: {file_url}")
            
            if not file_url:
                print("❌ [ERROR] 未提供合法的 URL")
                return jsonify({"status": "error", "message": "No URL provided"}), 400

            print(f"⬇️  [ACTION] 发起网络请求下载文件...")
            response = requests.get(file_url, timeout=30)
            response.raise_for_status()
            
            temp_dir = tempfile.gettempdir()
            raw_path = os.path.join(temp_dir, f"input_{os.getpid()}.doc")
            
            with open(raw_path, 'wb') as f:
                f.write(response.content)
            print(f"✅ [SUCCESS] 文件已下载到本地: {raw_path} (大小: {len(response.content)} bytes)")

            try:
                print(f"🔄 [ACTION] 正在 Windows 环境下调用原生 Word 转换: {raw_path}")
                
                abs_raw_path = os.path.abspath(raw_path)
                base_name = os.path.splitext(os.path.basename(abs_raw_path))[0]
                docx_path = os.path.join(temp_dir, f"{base_name}.docx")
                abs_docx_path = os.path.abspath(docx_path)

                pythoncom.CoInitialize()

                word = None
                doc = None
                try:
                    word = win32.Dispatch('Word.Application')
                    word.Visible = False
                    word.DisplayAlerts = False

                    doc = word.Documents.Open(abs_raw_path)
                    doc.SaveAs2(abs_docx_path, FileFormat=16)
                    
                    print(f"✅ [SUCCESS] 原生 Word 转换成功: {abs_docx_path}")

                finally:
                    if doc:
                        doc.Close(False)
                    if word:
                        word.Quit()
                    pythoncom.CoUninitialize()

                if not os.path.exists(abs_docx_path):
                    raise Exception("Word 运行结束但未生成 docx 文件")

            except Exception as e:
                print(f"❌ [DEBUG ERROR] Windows 原生转换失败: {str(e)}")
                return jsonify({"status": "error", "message": f"Windows Word conversion failed: {str(e)}"}), 500

            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            
            full_json_path = os.path.join(data_dir, 'full_parsed.json')
            citations_path = os.path.join(data_dir, 'reference_config.json')
            
            print(f"🔄 [ACTION] 开始提取 DOCX 内容为 JSON...")
            elements = parse_docx_to_json(docx_path, full_json_path, citations_path)
            
            if not elements:
                print("❌ [ERROR] DOCX 转换 JSON 失败或返回为空")
                return jsonify({'error': 'Failed to parse file'}), 500
            
            text_elements = extract_text_from_parsed_json(full_json_path,remove_md=remove_md)
            
            result = {"text_elements": text_elements}
            blocks_path = os.path.join(data_dir, 'parsed_blocks.json')
            with open(blocks_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            
            citations_count = 0
            if os.path.exists(citations_path):
                with open(citations_path, 'r', encoding='utf-8') as f:
                    citations_data = json.load(f)
                    citations_count = len(citations_data)
            
            print(f"📊 [INFO] 解析完成: 提取了 {len(text_elements)} 个文本块，全部元素共 {len(elements)} 个，引用 {citations_count} 个")
            print(f"💾 [SAVE] 简略文本块已保存至: {blocks_path}")
            print(f"💾 [SAVE] 完整 JSON 数据已保存至: {full_json_path}")
            if citations_count > 0:
                print(f"💾 [SAVE] 引用配置已保存至: {citations_path}")
            print("🏁 [API OUT] 请求处理成功返回 200")
            
            return jsonify({
                'status': 'success', 
                'message': 'File parsed successfully',
                'text_count': len(text_elements),
                'total_count': len(elements),
                'citations_count': citations_count
            }), 200
        except requests.exceptions.Timeout:
            print("❌ [NETWORK ERROR] 下载文件超时 (Timeout > 30s)！这通常是目标 URL 无法访问或防火墙拦截。")
            return jsonify({'error': 'File download timeout. Is the URL accessible from this server?'}), 504
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /parse-docx-file 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500
        finally:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
                print(f"🧹 [CLEANUP] 清理临时文件: {temp_path}")
            if converted_path and os.path.exists(converted_path):
                os.remove(converted_path)
                print(f"🧹 [CLEANUP] 清理转换文件: {converted_path}")

    @app.route('/identify-short-blocks', methods=['GET'])
    def identify_short_blocks_endpoint():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: GET /identify-short-blocks")
        try:
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            blocks_path = os.path.join(data_dir, 'parsed_blocks.json')
            
            if not os.path.exists(blocks_path):
                print(f"❌ [ERROR] 找不到之前解析的文本块文件: {blocks_path}")
                return jsonify({'error': 'No parsed blocks found'}), 404
            
            print(f"📂 [READ] 正在读取文件: {blocks_path}")
            with open(blocks_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            blocks = data.get("text_elements", data)
            print(f"🔄 [ACTION] 开始识别短文本块... (总输入块数: {len(blocks)})")
            
            short_blocks = identify_short_blocks(blocks)
            
            print(f"✅ [SUCCESS] 识别完毕，从 {len(blocks)} 个总块中找出了 {len(short_blocks)} 个短文本块")
            print("🏁 [API OUT] 请求处理成功返回 200")
            
            return jsonify({'short_blocks': short_blocks, "total_blocks": data})
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /identify-short-blocks 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/analyze-styles', methods=['POST'])
    def analyze_styles():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /analyze-styles")
        try:
            data = request.json or {}
            styles = data.get('styles', [])
            
            print(f"📦 [PAYLOAD] 接收到的 styles 数量: {len(styles)}")
            if len(styles) > 0:
                print(f"📄 [PREVIEW] 首个 style 预览: {json.dumps(styles[0], ensure_ascii=False)}")
            
            if not styles:
                print("❌ [ERROR] 未提供 styles 数据")
                return jsonify({'error': 'No styles provided'}), 400
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            output_path = os.path.join(data_dir, 'parsed_styles.json')
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(styles, f, ensure_ascii=False, indent=2)
            
            print(f"✅ [SUCCESS] 样式数据已保存至: {output_path}")
            print("🏁 [API OUT] 请求处理成功返回 200")
            
            return jsonify({'status': 'success', 'message': 'Styles saved successfully'}), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /analyze-styles 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500
    @app.route('/generate-document', methods=['POST'])
    def generate_document():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /generate-document")
        try:
            data = request.json or {}
            parsed_styles = data.get('parsed_styles', [])
            
            print(f"📦 [PAYLOAD] 接收到的 parsed_styles 数量: {len(parsed_styles)}")
            
            if not parsed_styles:
                print("❌ [ERROR] 未提供 parsed_styles 数据")
                return jsonify({'error': 'No parsed styles provided'}), 400
            
            downloads_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'download')
            os.makedirs(downloads_dir, exist_ok=True)
            
            # 优化：加上时间戳防止多个用户同时请求时文件被互相覆盖
            timestamp = int(time.time())
            output_filename = f'generated_document_{timestamp}.docx'
            output_path = os.path.join(downloads_dir, output_filename)
            
            print(f"🔄 [ACTION] 开始生成 DOCX 文档...")
            success = generate_docx_document(parsed_styles, output_path)
            
            if not success:
                print("❌ [ERROR] 文档生成核心逻辑 (generate_docx_document) 返回 False")
                return jsonify({'error': 'Failed to generate docx document'}), 500
            
            # 核心修改：生成 HTTP 下载链接
            # request.host_url 会自动获取当前服务器的地址 (如 http://127.0.0.1:5000/ 或公网 IP)
            download_url = f"{request.host_url.rstrip('/')}/download/{output_filename}"
            
            print(f"✅ [SUCCESS] DOCX 文档生成完毕，保存至本地: {output_path}")
            print(f"🔗 [LINK] 生成下载链接: {download_url}")
            print("🏁 [API OUT] 请求处理成功返回 200")
            
            return jsonify({
                'status': 'success', 
                'message': 'Document generated successfully', 
                'file_url': download_url,   # 返回给客户端的 HTTP 下载链接
                'local_path': output_path   # 保留本地路径方便调试查阅
            }), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /generate-document 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    # ================= 新增：专门用于提供文件下载的路由 =================
    @app.route('/download/<filename>', methods=['GET'])
    def download_file(filename):
        print(f"⬇️  [DOWNLOAD] 客户端请求下载文件: {filename}")
        downloads_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'download')
        
        if not os.path.exists(os.path.join(downloads_dir, filename)):
            print(f"❌ [ERROR] 请求的文件不存在: {filename}")
            return jsonify({'error': 'File not found'}), 404
            
        # send_from_directory 会安全地把服务器本地文件通过 HTTP 发送给客户端
        # as_attachment=True 表示强制浏览器下载，而不是直接在浏览器里打开
        return send_from_directory(downloads_dir, filename, as_attachment=True)

    @app.route('/backfill-styles', methods=['POST'])
    def backfill_styles_endpoint():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /backfill-styles")
        try:
            data = request.json or {}
            edited_elements = data.get('edited_elements', [])
            
            print(f"📦 [PAYLOAD] 接收到的 edited_elements 数量: {len(edited_elements)}")
            
            if not edited_elements:
                print("❌ [ERROR] 未提供 edited_elements 数据")
                return jsonify({'error': 'No edited elements provided'}), 400
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            full_json_path = os.path.join(data_dir, 'full_parsed.json')
            
            if not os.path.exists(full_json_path):
                print(f"❌ [ERROR] 找不到完整解析数据: {full_json_path}")
                return jsonify({'error': 'No full parsed data found. Please parse docx first.'}), 404
            
            print(f"📂 [READ] 读取 full_parsed.json...")
            with open(full_json_path, 'r', encoding='utf-8') as f:
                full_data = json.load(f)
            
            print(f"🔄 [ACTION] 正在回填样式 (backfill_styles)...")
            from utils.docx_style_backfill import backfill_styles
            updated_data = backfill_styles(edited_elements, full_data)
            
            output_path = os.path.join(data_dir, 'backfilled_styles.json')
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(updated_data, f, ensure_ascii=False, indent=2)
            
            print(f"✅ [SUCCESS] 回填样式完毕，已保存至: {output_path}")
            print("🏁 [API OUT] 请求处理成功返回 200")
            
            return jsonify({
                'status': 'success', 
                'message': 'Styles backfilled successfully',
                'data': updated_data
            }), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /backfill-styles 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500
    @app.route('/restore-document', methods=['POST'])
    def restore_document():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /restore-document")
        try:
            
            print("⚠️ [WARN] 未收到 data 参数，尝试读取本地 backfilled_styles.json...")
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            json_path = os.path.join(data_dir, 'backfilled_styles.json')
            if os.path.exists(json_path):
                with open(json_path, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                print(f"📂 [READ] 成功从本地读取数据，长度: {len(json_data)}")
            
            if not json_data:
                print("❌ [ERROR] 既没有传入数据，也没有找到本地数据")
                return jsonify({'error': 'No data provided or found'}), 400
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            citations_path = os.path.join(data_dir, 'reference_config.json')
            
            citations = None
            if os.path.exists(citations_path):
                with open(citations_path, 'r', encoding='utf-8') as f:
                    citations = json.load(f)
                print(f"📚 [CITATIONS] 读取到 {len(citations)} 个引用配置")
            
            downloads_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'download')
            os.makedirs(downloads_dir, exist_ok=True)
            
            timestamp = int(time.time())
            filename = f'restored_document_{timestamp}.docx'
            output_path = os.path.join(downloads_dir, filename)
            
            temp_json_path = os.path.join(downloads_dir, f'temp_restore_{timestamp}.json')
            
            with open(temp_json_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            
            print(f"🔄 [ACTION] 开始根据 JSON 还原文档 (restore_docx_from_json)...")
            success = restore_docx_from_json(temp_json_path, output_path, citations_path if citations else None)
            
            if os.path.exists(temp_json_path):
                os.remove(temp_json_path)
                print(f"🧹 [CLEANUP] 清理临时文件: {temp_json_path}")
            
            if not success:
                print("❌ [ERROR] 还原文档失败 (restore_docx_from_json 返回 False)")
                return jsonify({'error': 'Failed to restore document'}), 500
            
            download_url = f"/download/{filename}"
            
            print(f"✅ [SUCCESS] 还原文档成功，保存至: {output_path}")
            print(f"🔗 [LINK] 成功生成下载链接: {download_url}")
            print("🏁 [API OUT] 请求处理成功返回 200")
            
            return jsonify({
                'status': 'success', 
                'message': 'Document restored successfully',
                'file_url': download_url,
                'local_path': output_path,
                'citations_restored': len(citations) if citations else 0
            }), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /restore-document 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/parse-text', methods=['POST'])
    def parse_text():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /parse-text")
        try:
            data = request.json or {}
            text = data.get('text', '')
            remove_md = data.get('remove_markdown', True)
            
            print(f"📦 [PAYLOAD] 接收到文本内容，长度: {len(text)} 字符")
            print(f"⚙️  [PARSE] 是否移除 Markdown: {remove_md}")
            
            if not text:
                print("❌ [ERROR] 未提供 text 内容")
                return jsonify({'error': 'No text provided'}), 400
            
            print(f"🔄 [ACTION] 正在解析纯文本...")
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            
            full_json_path = os.path.join(data_dir, 'full_parsed.json')
            elements = parse_text_to_json(text, full_json_path, remove_md)
            
            if not elements:
                print("❌ [ERROR] 解析纯文本失败")
                return jsonify({'error': 'Failed to parse text'}), 500
            
            text_elements = []
            for elem in elements:
                text_elements.append({
                    "id": elem.get("id"),
                    "content": elem.get("content")
                })
            
            result = {"text_elements": text_elements}
            blocks_path = os.path.join(data_dir, 'parsed_blocks.json')
            with open(blocks_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            
            print(f"✅ [SUCCESS] 解析完毕，共生成 {len(text_elements)} 个元素")
            print(f"💾 [SAVE] 简略文本块保存至: {blocks_path}")
            print(f"💾 [SAVE] 完整 JSON 保存至: {full_json_path}")
            print("🏁 [API OUT] 请求处理成功返回 200")
            
            return jsonify({
                'status': 'success',
                'message': 'Text parsed successfully',
                'text_count': len(text_elements),
                'total_count': len(elements)
            }), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /parse-text 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/text-to-docx', methods=['POST'])
    def text_to_docx():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /text-to-docx")
        try:
            data = request.json or {}
            text = data.get('text', '')
            remove_md = data.get('remove_markdown', True)
            styles = data.get('styles', [])
            
            print(f"📦 [PAYLOAD] 文本长度: {len(text)} 字符, 样式规则数量: {len(styles)}")
            
            if not text:
                print("❌ [ERROR] 未提供 text 内容")
                return jsonify({'error': 'No text provided'}), 400
            
            print(f"🔄 [ACTION] 将文本解析为基础 JSON...")
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            
            full_json_path = os.path.join(data_dir, 'full_parsed.json')
            elements = parse_text_to_json(text, full_json_path, remove_md)
            
            if not elements:
                print("❌ [ERROR] 解析纯文本失败")
                return jsonify({'error': 'Failed to parse text'}), 500
            
            if styles:
                print(f"🔄 [ACTION] 正在应用 {len(styles)} 条样式规则...")
                id_to_style = {s.get('id'): s for s in styles}
                for elem in elements:
                    if elem.get('id') in id_to_style:
                        style_info = id_to_style[elem.get('id')]
                        new_type = style_info.get('type', 'body')
                        if new_type.startswith('heading'):
                            elem['type'] = new_type
                        elif new_type == 'normal':
                            elem['type'] = 'body'
            
            output_json_path = os.path.join(data_dir, 'text_styled.json')
            with open(output_json_path, 'w', encoding='utf-8') as f:
                json.dump(elements, f, ensure_ascii=False, indent=2)
            
            downloads_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'download')
            os.makedirs(downloads_dir, exist_ok=True)
            output_path = os.path.join(downloads_dir, 'text_document.docx')
            
            print(f"🔄 [ACTION] 正在生成 DOCX 文件...")
            success = restore_docx_from_json(output_json_path, output_path)
            
            if not success:
                print("❌ [ERROR] 还原文档失败 (restore_docx_from_json 返回 False)")
                return jsonify({'error': 'Failed to generate document'}), 500
            
            print(f"✅ [SUCCESS] 从纯文本生成文档成功，保存至: {output_path}")
            print("🏁 [API OUT] 请求处理成功返回 200")
            
            return jsonify({
                'status': 'success',
                'message': 'Document generated successfully',
                'file_path': output_path,
                'element_count': len(elements)
            }), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /text-to-docx 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/detect-formulas', methods=['GET'])
    def detect_formulas():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: GET /detect-formulas")
        try:
            from core.formula.detector import detect_formula_blocks
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            json_path = os.path.join(data_dir, 'parsed_blocks.json')

            if not os.path.exists(json_path):
                print(f"❌ [ERROR] 找不到 parsed_blocks.json: {json_path}")
                return jsonify({'error': 'No parsed_blocks.json found. Please run backfill-styles first.'}), 404

            with open(json_path, 'r', encoding='utf-8') as f:
                elements = json.load(f)

            suspected = detect_formula_blocks(elements)
            print(f"✅ [SUCCESS] 从 {len(elements)} 个元素中检测到 {len(suspected)} 个疑似公式块")
            print("🏁 [API OUT] 请求处理成功返回 200")

            return jsonify({'suspected_formulas': suspected, 'count': len(suspected)}), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /detect-formulas 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/convert-formulas', methods=['POST'])
    def convert_formulas():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /convert-formulas")
        try:
            from core.formula.converter import convert_formula_list
            data = request.json or {}
            formula_items = data.get('formulas', [])

            print(f"📦 [PAYLOAD] 接收到 {len(formula_items)} 个公式条目")

            if not formula_items:
                print("❌ [ERROR] 未提供 formulas 数据")
                return jsonify({'error': 'No formulas provided'}), 400

            results = convert_formula_list(formula_items)

            failed = [r for r in results if r.get('error')]
            print(f"✅ [SUCCESS] 转换完成: {len(results) - len(failed)} 成功, {len(failed)} 失败")
            print("🏁 [API OUT] 请求处理成功返回 200")

            return jsonify({'results': results, 'total': len(results), 'failed': len(failed)}), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /convert-formulas 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/process-formulas', methods=['GET'])
    def process_formulas():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: GET /process-formulas")
        try:
            from core.formula.detector import detect_formula_blocks, merge_formula_blocks
            from core.formula.converter import convert_formula_list
            from core.formula.models import FormulaListResponse
            from utils.base_agent import call_structured

            # Step 1: 读取数据
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            json_path = os.path.join(data_dir, 'parsed_blocks.json')

            if not os.path.exists(json_path):
                print(f"❌ [ERROR] 找不到 parsed_blocks.json")
                return jsonify({'error': 'No parsed_blocks.json found. Please run backfill-parsed_blocks first.'}), 404

            with open(json_path, 'r', encoding='utf-8') as f:
                elements = json.load(f).get("text_elements",[])

            # Step 2: 合并 $$ 块内的多行片段，再规则检测疑似公式
            elements = merge_formula_blocks(elements)
            suspected = detect_formula_blocks(elements)
            print(f"🔍 [DETECT] 从 {len(elements)} 个元素中检测到 {len(suspected)} 个疑似公式块")

            if not suspected:
                print("✅ [SUCCESS] 未检测到公式，直接返回")
                return jsonify({'results': []}), 200

            # Step 3: LLM 提取确认公式（分批处理，每批 BATCH_SIZE 条）
            BATCH_SIZE = 7
            system_prompt = (
                "你是一个数学公式提取器。你会收到一组文本元素（每个元素含 id 和 content），"
                "任务是从中识别并提取数学公式，将其转换为标准 LaTeX 格式。\n\n"
                "判断规则：\n"
                "- 明确的数学表达式（含变量、运算符、等式）\n"
                "- LaTeX 语法片段（如 \\frac、\\sum、\\int 等）\n"
                "- 含数学符号的文本（∑ ∫ √ ± × ÷ ∞ 等）\n"
                "- 上下标结构（x² a₁ 等）\n\n"
                "不视为公式：纯自然语言描述、单独的数字或百分比、无意义符号。\n\n"
                "提取规则：\n"
                "- 行内公式：text_before 填公式前文本，latex_formula 填公式，text_after 填公式后文本，label 为空\n"
                "- 块级公式：text_before 和 text_after 为空，label 填编号（如 (4-1)）或空，latex_formula 填公式本体\n"
                "- 若一个元素含多个行内公式，针对每个公式分别输出一条 FormulaItem（id 相同）；"
                "  除最后一条外 text_after 留空，最后一条 text_after 填最后公式之后的全部剩余文本\n"
                "- 不含公式的元素不出现在输出中\n"
                "- latex_formula 只写公式本体，不加 $ 符号"
            )

            all_formulas = []
            total_batches = (len(suspected) + BATCH_SIZE - 1) // BATCH_SIZE
            for i in range(0, len(suspected), BATCH_SIZE):
                batch = suspected[i:i + BATCH_SIZE]
                user_prompt = json.dumps(batch, ensure_ascii=False)
                batch_response = call_structured(system_prompt, user_prompt, FormulaListResponse)
                all_formulas.extend(batch_response.formulas)
                print(f"🤖 [LLM] 批次 {i // BATCH_SIZE + 1}/{total_batches}：提取 {len(batch_response.formulas)} 个公式")

            print(f"🤖 [LLM] 共确认提取 {len(all_formulas)} 个公式")

            if not all_formulas:
                print("✅ [SUCCESS] LLM 确认无有效公式")
                return jsonify({'results': []}), 200

            # Step 4: latex → omath
            formula_dicts = [f.model_dump() for f in all_formulas]
            results = convert_formula_list(formula_dicts)

            failed = [r for r in results if r.get('error')]
            print(f"✅ [SUCCESS] omath 转换完成，{len(results) - len(failed)} 成功，{len(failed)} 失败")
            print("🏁 [API OUT] 请求处理成功返回 200")

            return jsonify({'results': results}), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /process-formulas 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/detect-tables', methods=['GET'])
    def detect_tables():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: GET /detect-tables")
        try:
            from core.table.detector import detect_table_blocks
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            json_path = os.path.join(data_dir, 'parsed_blocks.json')

            if not os.path.exists(json_path):
                print(f"❌ [ERROR] 找不到 parsed_blocks.json: {json_path}")
                return jsonify({'error': 'No parsed_blocks.json found. Please parse a file first.'}), 404

            with open(json_path, 'r', encoding='utf-8') as f:
                elements = json.load(f).get("text_elements", [])

            suspected = detect_table_blocks(elements)
            print(f"✅ [SUCCESS] 从 {len(elements)} 个元素中检测到 {len(suspected)} 个疑似表格块")
            print("🏁 [API OUT] 请求处理成功返回 200")

            return jsonify({'suspected_tables': suspected, 'count': len(suspected)}), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /detect-tables 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/process-tables', methods=['GET'])
    def process_tables():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: GET /process-tables")
        try:
            from core.table.detector import detect_table_blocks, group_table_blocks, is_table_title
            from core.table.extractor import extract_table
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            json_path = os.path.join(data_dir, 'parsed_blocks.json')

            if not os.path.exists(json_path):
                print(f"❌ [ERROR] 找不到 parsed_blocks.json")
                return jsonify({'error': 'No parsed_blocks.json found. Please parse a file first.'}), 404

            with open(json_path, 'r', encoding='utf-8') as f:
                elements = json.load(f).get("text_elements", [])

            suspected = detect_table_blocks(elements)
            print(f"🔍 [DETECT] 从 {len(elements)} 个元素中检测到 {len(suspected)} 个疑似表格行")

            if not suspected:
                print("✅ [SUCCESS] 未检测到表格，直接返回")
                return jsonify({'results': []}), 200

            groups = group_table_blocks(elements, suspected)
            print(f"📊 [GROUP] 合并为 {len(groups)} 个表格组")

            id_to_pos = {elem["id"]: i for i, elem in enumerate(elements)}

            results = []
            for group in groups:
                ids = [item['id'] for item in group]
                combined = "\n".join(item['content'] for item in group)

                # 感知前置表题
                existing_title = None
                first_pos = id_to_pos.get(ids[0], -1)
                if first_pos > 0:
                    preceding = elements[first_pos - 1]
                    if is_table_title(preceding.get("content", "")):
                        existing_title = preceding["content"]
                        print(f"📌 [TITLE] 检测到已有表题: {existing_title}")

                print(f"🤖 [LLM] 正在处理表格组 {ids[0]}~{ids[-1]}（{len(ids)} 行）...")
                blocks = extract_table(combined, existing_title=existing_title)
                if blocks is None:
                    print(f"⚠️  [SKIP] 模型判断该组非表格，跳过")
                    continue
                results.append({'ids': ids, 'blocks': blocks})

            print(f"✅ [SUCCESS] 表格提取完成，共处理 {len(results)} 个表格")
            print("🏁 [API OUT] 请求处理成功返回 200")

            return jsonify({'results': results}), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /process-tables 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/backfill-tables', methods=['POST'])
    def backfill_tables():
        print("\n" + "="*50)
        print(f"🌐 [API IN] 收到请求: POST /backfill-tables")
        try:
            data = request.json or {}
            results = data.get('results', [])

            if not results:
                return jsonify({'error': 'No results provided'}), 400

            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            json_path = os.path.join(data_dir, 'backfilled_styles.json')

            if not os.path.exists(json_path):
                return jsonify({'error': 'No backfilled_styles.json found. Please run backfill-styles first.'}), 404

            with open(json_path, 'r', encoding='utf-8') as f:
                elements = json.load(f)

            # ids[0] 替换为 blocks，ids[1:] 整行删除
            first_id_to_blocks = {}
            ids_to_drop = set()
            for item in results:
                ids = item.get('ids', [item.get('id')])  # 兼容旧单 id 格式
                if not ids:
                    continue
                first_id_to_blocks[ids[0]] = item['blocks']
                for extra_id in ids[1:]:
                    ids_to_drop.add(extra_id)

            new_elements = []
            replaced = 0
            for elem in elements:
                elem_id = elem.get('id')
                if elem_id in ids_to_drop:
                    continue
                elif elem_id in first_id_to_blocks:
                    new_elements.extend(first_id_to_blocks[elem_id])
                    replaced += 1
                else:
                    new_elements.append(elem)

            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(new_elements, f, ensure_ascii=False, indent=2)

            print(f"✅ [SUCCESS] 表格回填完成，替换了 {replaced} 个表格，文档现共 {len(new_elements)} 个块")
            print("🏁 [API OUT] 请求处理成功返回 200")

            return jsonify({'replaced': replaced, 'total_blocks': len(new_elements)}), 200
        except Exception as e:
            print(f"❌ [CRITICAL ERROR] /backfill-tables 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500
        
    @app.route('/process-short-blocks', methods=['GET'])                                                                           
    def process_short_blocks():                                                                                                           
        print("\n" + "="*50)                                                                                                       
        print(f"🌐 [API IN] 收到请求: GET /process-short-blocks")                                                                  
        try:                                                                                                                       
            from core.short_block.classifier import classify_short_blocks                                                          
                                                                                                                                   
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')                           
            json_path = os.path.join(data_dir, 'parsed_blocks.json')                                                               
                                                                                                                                   
            if not os.path.exists(json_path):                                                                                      
                print(f"❌ [ERROR] 找不到 parsed_blocks.json: {json_path}")                                                        
                return jsonify({'error': 'No parsed_blocks.json found. Please parse a file first.'}), 404                          
                                                                                                                                   
            with open(json_path, 'r', encoding='utf-8') as f:                                                                      
                blocks = json.load(f).get("text_elements", [])                                                                     
                                                                                                                                   
            short_blocks = identify_short_blocks(blocks)                                                                           
            print(f"🔍 [DETECT] 从 {len(blocks)} 个元素中识别到 {len(short_blocks)} 个短文本块")                                   
                                                                                                                                   
            if not short_blocks:                                                                                                   
                print("✅ [SUCCESS] 未检测到短文本块，直接返回")                                                                   
                return jsonify({'results': []}), 200                                                                               
                                                                                                                                   
            results = classify_short_blocks(short_blocks)                                                                          
            print(f"🤖 [LLM] 共分类 {len(results)} 个短文本块")                                                                    
            print("🏁 [API OUT] 请求处理成功返回 200")                                                                             
                                                                                                                                   
            return jsonify({'results': results}), 200                                                                              
        except Exception as e:                                                                                                     
            print(f"❌ [CRITICAL ERROR] /process-short-blocks 运行异常: {e}")                                                      
            traceback.print_exc()                                                                                                  
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500                                            
                                                                                                                                   
    @app.route('/backfill-short-blocks', methods=['POST'])                                                                         
    def backfill_short_blocks():                                                                                                   
        print("\n" + "="*50)                                                                                                       
        print(f"🌐 [API IN] 收到请求: POST /backfill-short-blocks")                                                                
        try:                                                                                                                       
            data = request.json or {}                                                                                              
            results = data.get('results', [])                                                                                      
                                                                                                                                   
            if not results:                                                                                                        
                return jsonify({'error': 'No results provided'}), 400                                                              
                                                                                                                                   
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')                           
            full_json_path = os.path.join(data_dir, 'full_parsed.json')                                                            
                                                                                                                                   
            if not os.path.exists(full_json_path):                                                                                 
                print(f"❌ [ERROR] 找不到 full_parsed.json: {full_json_path}")                                                     
                return jsonify({'error': 'No full_parsed.json found. Please parse a file first.'}), 404                            
                                                                                                                                   
            with open(full_json_path, 'r', encoding='utf-8') as f:                                                                 
                full_data = json.load(f)                                                                                           
                                                                                                                                   
            from utils.docx_style_backfill import backfill_styles                                                                  
            updated_data = backfill_styles(results, full_data)                                                                     
                                                                                                                                   
            output_path = os.path.join(data_dir, 'backfilled_styles.json')                                                         
            with open(output_path, 'w', encoding='utf-8') as f:                                                                    
                json.dump(updated_data, f, ensure_ascii=False, indent=2)                                                           
                                                                                                                                   
            print(f"[SUCCESS] 短文本块回填完成，已保存至:: {output_path}")
            print("[API OUT]  请求处理成功返回 200")

            return jsonify({
                'status': 'success',
            }), 200
        except Exception as e:                                                                                                     
            print(f"❌ [CRITICAL ERROR] /backfill-short-blocks 运行异常: {e}")                                                     
            traceback.print_exc()                                                                                                  
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500                                            
                                                                                            
    @app.route('/backfill-formulas', methods=['POST'])
    def backfill_formulas():
        print("\n" + "="*50)
        print(f"[API IN] POST /backfill-formulas")
        try:
            data = request.json or {}
            results = data.get('results', [])

            if not results:
                return jsonify({'error': 'No results provided'}), 400

            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            json_path = os.path.join(data_dir, 'backfilled_styles.json')

            if not os.path.exists(json_path):
                return jsonify({'error': 'No backfilled_styles.json found'}), 404

            with open(json_path, 'r', encoding='utf-8') as f:
                elements = json.load(f)

            # 合并 $$ 块内的片段，删掉 elem_39/40/41 这类残留碎片
            from core.formula.detector import merge_formula_blocks
            elements = merge_formula_blocks(elements)

            id_to_elem = {e.get('id'): e for e in elements}

            # 按 id 分组，同一元素的多条公式聚合在一起
            from collections import defaultdict
            id_to_items = defaultdict(list)
            for item in results:
                elem_id = item.get('id')
                if elem_id and item.get('omath'):
                    id_to_items[elem_id].append(item)

            updated = 0
            for elem_id, items in id_to_items.items():
                if elem_id not in id_to_elem:
                    continue
                if len(items) == 1:
                    item = items[0]
                    text_before = item.get('text_before', '')
                    text_after = item.get('text_after', '')
                    formula_type = 'formula_inline' if (text_before or text_after) else 'formula_block'
                    id_to_elem[elem_id]['type'] = formula_type
                    id_to_elem[elem_id]['formula'] = {
                        'text_before': text_before,
                        'omath': item.get('omath', ''),
                        'text_after': text_after,
                        'label': item.get('label', ''),
                    }
                else:
                    id_to_elem[elem_id]['type'] = 'formula_inline_multi'
                    id_to_elem[elem_id]['formula_segments'] = [
                        {
                            'text_before': it.get('text_before', ''),
                            'omath': it.get('omath', ''),
                            'text_after': it.get('text_after', ''),
                        }
                        for it in items
                    ]
                updated += 1

            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(elements, f, ensure_ascii=False, indent=2)

            print(f"[SUCCESS] 回填完成，更新了 {updated} 个公式元素")
            return jsonify({'updated': updated}), 200
        except Exception as e:
            print(f"[CRITICAL ERROR] /backfill-formulas 运行异常: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500