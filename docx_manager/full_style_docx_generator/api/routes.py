from flask import request, jsonify
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

    @app.route('/parse-txt-file', methods=['POST'])
    def parse_txt_file():
        temp_path = None
        try:
            data = request.json
            url = data.get('url', '')
            remove_md = data.get('remove_markdown', True)
            
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
            
            result = parse_file(temp_path, remove_md=remove_md)
            
            if not result:
                return jsonify({'error': 'Failed to parse file'}), 500
            
            blocks = result.get("text_elements", [])
            print(f"Parsed {len(blocks)} blocks from file")
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            output_path = os.path.join(data_dir, 'parsed_blocks.json')
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            
            print(f"Saved parsed blocks to: {output_path}")
            
            return jsonify({'status': 'success', 'message': 'File parsed and blocks saved', 'text_count': len(blocks)}), 200
        except Exception as e:
            print(f"Error parsing txt file: {e}")
            traceback.print_exc()
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
            doc.SaveAs2(abs_docx_path, FileFormat=16)
            
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
            
            response = requests.get(file_url, timeout=30)
            temp_dir = tempfile.gettempdir()
            raw_path = os.path.join(temp_dir, f"input_{os.getpid()}.doc")
            
            with open(raw_path, 'wb') as f:
                f.write(response.content)
            print(f"[DEBUG] 文件已下载到: {raw_path}")

            try:
                print(f"[DEBUG] 正在 Windows 环境下调用原生 Word 转换: {raw_path}")
                
                abs_raw_path = os.path.abspath(raw_path)
                base_name = os.path.splitext(os.path.basename(abs_raw_path))[0]
                docx_path = os.path.join(temp_dir, f"{base_name}.docx")
                abs_docx_path = os.path.abspath(docx_path)

                pythoncom.CoInitialize()

                word = None
                doc = None
                try:
                    word = win32.gencache.EnsureDispatch('Word.Application')
                    word.Visible = False
                    word.DisplayAlerts = False

                    doc = word.Documents.Open(abs_raw_path)
                    doc.SaveAs2(abs_docx_path, FileFormat=16)
                    
                    print(f"[DEBUG] 原生 Word 转换成功: {abs_docx_path}")

                finally:
                    if doc:
                        doc.Close(False)
                    if word:
                        word.Quit()
                    pythoncom.CoUninitialize()

                if not os.path.exists(abs_docx_path):
                    raise Exception("Word 运行结束但未生成 docx 文件")

            except Exception as e:
                print(f"[DEBUG ERROR] Windows 原生转换失败: {str(e)}")
                return jsonify({"status": "error", "message": f"Windows Word conversion failed: {str(e)}"}), 500

            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            
            full_json_path = os.path.join(data_dir, 'full_parsed.json')
            elements = parse_docx_to_json(docx_path, full_json_path)
            
            if not elements:
                return jsonify({'error': 'Failed to parse file'}), 500
            
            text_elements = extract_text_from_parsed_json(full_json_path)
            
            result = {"text_elements": text_elements}
            blocks_path = os.path.join(data_dir, 'parsed_blocks.json')
            with open(blocks_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            
            print(f"Parsed {len(text_elements)} text elements, saved to {blocks_path}")
            print(f"Full parsed data saved to {full_json_path}")
            
            return jsonify({
                'status': 'success', 
                'message': 'File parsed successfully',
                'text_count': len(text_elements),
                'total_count': len(elements)
            }), 200
        except Exception as e:
            print(f"Error parsing docx file: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500
        finally:
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
                data = json.load(f)
            
            blocks = data.get("text_elements", data)
            
            short_blocks = identify_short_blocks(blocks)
            
            print(f"Identified {len(short_blocks)} short blocks from {len(blocks)} total blocks")
            
            return jsonify({'short_blocks': short_blocks, "total_blocks": data})
        except Exception as e:
            print(f"Error identifying short blocks: {e}")
            traceback.print_exc()
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
            traceback.print_exc()
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
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/backfill-styles', methods=['POST'])
    def backfill_styles_endpoint():
        try:
            data = request.json
            edited_elements = data.get('edited_elements', [])
            
            if not edited_elements:
                return jsonify({'error': 'No edited elements provided'}), 400
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            full_json_path = os.path.join(data_dir, 'full_parsed.json')
            
            if not os.path.exists(full_json_path):
                return jsonify({'error': 'No full parsed data found. Please parse docx first.'}), 404
            
            with open(full_json_path, 'r', encoding='utf-8') as f:
                full_data = json.load(f)
            
            from utils.docx_style_backfill import backfill_styles
            updated_data = backfill_styles(edited_elements, full_data)
            
            output_path = os.path.join(data_dir, 'backfilled_styles.json')
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(updated_data, f, ensure_ascii=False, indent=2)
            
            print(f"Backfilled styles saved to: {output_path}")
            
            return jsonify({
                'status': 'success', 
                'message': 'Styles backfilled successfully',
                'data': updated_data
            }), 200
        except Exception as e:
            print(f"Error backfilling styles: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/restore-document', methods=['POST'])
    def restore_document():
        try:
            data = request.json
            json_data = data.get('data', [])
            
            if not json_data:
                data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
                json_path = os.path.join(data_dir, 'backfilled_styles.json')
                if os.path.exists(json_path):
                    with open(json_path, 'r', encoding='utf-8') as f:
                        json_data = json.load(f)
            
            if not json_data:
                return jsonify({'error': 'No data provided or found'}), 400
            
            downloads_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'download')
            os.makedirs(downloads_dir, exist_ok=True)
            
            output_path = os.path.join(downloads_dir, 'restored_document.docx')
            
            temp_json_path = os.path.join(downloads_dir, 'temp_restore.json')
            with open(temp_json_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            
            success = restore_docx_from_json(temp_json_path, output_path)
            
            if os.path.exists(temp_json_path):
                os.remove(temp_json_path)
            
            if not success:
                return jsonify({'error': 'Failed to restore document'}), 500
            
            print(f"Restored document saved to: {output_path}")
            
            return jsonify({
                'status': 'success', 
                'message': 'Document restored successfully',
                'file_path': output_path
            }), 200
        except Exception as e:
            print(f"Error restoring document: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/parse-text', methods=['POST'])
    def parse_text():
        try:
            data = request.json
            text = data.get('text', '')
            remove_md = data.get('remove_markdown', True)
            
            if not text:
                return jsonify({'error': 'No text provided'}), 400
            
            print(f"Parsing text ({len(text)} chars), remove_markdown={remove_md}")
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            
            full_json_path = os.path.join(data_dir, 'text_parsed.json')
            elements = parse_text_to_json(text, full_json_path, remove_md)
            
            if not elements:
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
            
            print(f"Parsed {len(text_elements)} text elements, saved to {blocks_path}")
            print(f"Full parsed data saved to {full_json_path}")
            
            return jsonify({
                'status': 'success',
                'message': 'Text parsed successfully',
                'text_count': len(text_elements),
                'total_count': len(elements)
            }), 200
        except Exception as e:
            print(f"Error parsing text: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    @app.route('/text-to-docx', methods=['POST'])
    def text_to_docx():
        try:
            data = request.json
            text = data.get('text', '')
            remove_md = data.get('remove_markdown', True)
            styles = data.get('styles', [])
            
            if not text:
                return jsonify({'error': 'No text provided'}), 400
            
            print(f"Processing text to docx ({len(text)} chars)")
            
            data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'data')
            os.makedirs(data_dir, exist_ok=True)
            
            full_json_path = os.path.join(data_dir, 'text_parsed.json')
            elements = parse_text_to_json(text, full_json_path, remove_md)
            
            if not elements:
                return jsonify({'error': 'Failed to parse text'}), 500
            
            if styles:
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
            
            success = restore_docx_from_json(output_json_path, output_path)
            
            if not success:
                return jsonify({'error': 'Failed to generate document'}), 500
            
            print(f"Document generated from text: {output_path}")
            
            return jsonify({
                'status': 'success',
                'message': 'Document generated successfully',
                'file_path': output_path,
                'element_count': len(elements)
            }), 200
        except Exception as e:
            print(f"Error processing text to docx: {e}")
            traceback.print_exc()
            return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500
