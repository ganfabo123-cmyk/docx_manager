"""
wps_com/header.py — COM-based header manipulation for WPS/Word

公开函数：
    clear_all_headers   清空所有节的页眉
    set_all_headers     将所有节的页眉设置为固定文字

依赖：pip install pywin32
"""
from __future__ import annotations

import os
import sys

_WD_HEADER_FOOTER_PRIMARY    = 1
_WD_HEADER_FOOTER_FIRST_PAGE = 2
_WD_HEADER_FOOTER_EVEN_PAGES = 3


def _get_app(visible: bool = False):
    try:
        import win32com.client as wc
        import pythoncom
    except ImportError:
        sys.exit('[ERROR] pywin32 未安装: pip install pywin32')
    pythoncom.CoInitialize()
    for prog_id in ['Kwps.Application', 'Word.Application']:
        try:
            app = wc.Dispatch(prog_id)
            app.Visible = visible
            print(f'[COM] 连接: {prog_id}')
            return app
        except Exception:
            continue
    sys.exit('[ERROR] 无法连接 WPS 或 Word，请确认已安装')


def _set_header_text(section, text: str) -> None:
    for hf_type in (_WD_HEADER_FOOTER_PRIMARY, _WD_HEADER_FOOTER_EVEN_PAGES,
                    _WD_HEADER_FOOTER_FIRST_PAGE):
        try:
            section.Headers(hf_type).Range.Text = text
        except Exception:
            pass


def clear_all_headers(docx_path: str, visible: bool = False) -> None:
    """清空文档所有节的页眉。"""
    import pythoncom
    abs_path = os.path.abspath(docx_path)
    app = _get_app(visible)
    try:
        doc = app.Documents.Open(abs_path)
        for i in range(1, doc.Sections.Count + 1):
            _set_header_text(doc.Sections(i), "")
        doc.Save()
        doc.Close()
        print(f'[DONE] 页眉已清空: {abs_path}')
    except Exception as e:
        print(f'[ERROR] {e}', file=sys.stderr)
        try:
            doc.Close(False)
        except Exception:
            pass
        raise
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def set_all_headers(docx_path: str, text: str, visible: bool = False) -> None:
    """将文档所有节的页眉设置为 text。"""
    import pythoncom
    abs_path = os.path.abspath(docx_path)
    app = _get_app(visible)
    try:
        doc = app.Documents.Open(abs_path)
        for i in range(1, doc.Sections.Count + 1):
            _set_header_text(doc.Sections(i), text)
        doc.Save()
        doc.Close()
        print(f'[DONE] 页眉已设置: {abs_path}')
    except Exception as e:
        print(f'[ERROR] {e}', file=sys.stderr)
        try:
            doc.Close(False)
        except Exception:
            pass
        raise
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
