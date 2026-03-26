import traceback
import os
from typing import Optional, List, Dict, Any

try:
    from docx import Document
except ImportError:
    print("python-docx not installed, docx parsing will be disabled")

def parse_txt_file(file_path: str) -> Optional[List[Dict[str, Any]]]:
    """
    Parse text from txt file into text blocks
    """
    try:
        blocks = []
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            for i, line in enumerate(content.split('\n'), 1):
                if line.strip():
                    blocks.append({
                        'id': i,
                        'content': line.strip()
                    })
        return blocks
    except Exception as e:
        print(f"Error parsing txt file: {e}")
        print(traceback.format_exc())
        return None

def parse_docx_file(file_path: str) -> Optional[List[Dict[str, Any]]]:
    """
    Parse text from docx file into text blocks
    """
    try:
        if not os.path.exists(file_path):
            print(f"File does not exist: {file_path}")
            return None
        
        file_size = os.path.getsize(file_path)
        print(f"File size: {file_size} bytes")
        
        if file_size == 0:
            print(f"File is empty: {file_path}")
            return None
        
        with open(file_path, 'rb') as f:
            header = f.read(4)
            print(f"File header (hex): {header.hex()}")
            
            if header[:4] == b'PK\x03\x04':
                print("File appears to be a valid ZIP/DOCX format")
            elif header[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                print("File appears to be a DOC format (OLE2), not DOCX")
                print("Error: This is a .doc file, not .docx. Please convert it to .docx format first.")
                return None
            else:
                print(f"Unknown file format. Header: {header}")
        
        blocks = []
        doc = Document(file_path)
        
        for i, paragraph in enumerate(doc.paragraphs, 1):
            if paragraph.text.strip():
                blocks.append({
                    'id': i,
                    'content': paragraph.text.strip()
                })
        
        print(f"Successfully parsed {len(blocks)} blocks from docx file")
        return blocks
        
    except Exception as e:
        print(f"Error parsing docx file: {e}")
        print(traceback.format_exc())
        return None

def parse_file(file_path: str) -> Optional[List[Dict[str, Any]]]:
    """
    Parse text from file (supports txt and docx) into text blocks
    """
    try:
        if not os.path.exists(file_path):
            print(f"File does not exist: {file_path}")
            return None
            
        if file_path.endswith('.txt'):
            return parse_txt_file(file_path)
        elif file_path.endswith('.docx'):
            return parse_docx_file(file_path)
        else:
            print(f"Unsupported file type: {file_path}")
            return None
    except Exception as e:
        print(f"Error parsing file: {e}")
        print(traceback.format_exc())
        return None