"""
wps_com/insert_image.py — COM-based image insertion for WPS/Word

公开函数：
    insert_n_images_one_col   单列图片插入
    insert_n_images_two_col   双列图片插入

依赖：pip install pywin32
"""
from __future__ import annotations

import os
import re
import sys

_CM_TO_PT = 28.35

_WD_ALIGN_CENTER   = 1
_WD_COLLAPSE_START = 1
_WD_COLLAPSE_END   = 0

LABEL_GAP_SPACES = 33   # 双列标签行 (a) 与 (b) 之间的空格数
MAX_LINE_CHARS   = 32   # 子图题行最大字符数


# ── 工具 ────────────────────────────────────────────────────────────────────────

def _cm(cm: float) -> float:
    return cm * _CM_TO_PT


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


def _find_anchor_para(doc, anchor_text: str) -> int:
    """扫描全文段落，返回包含 anchor_text 的段落的 1-based 全局序号。"""
    total = doc.Paragraphs.Count
    for i in range(1, total + 1):
        if anchor_text in doc.Paragraphs(i).Range.Text:
            return i
    raise RuntimeError(f'未找到锚定文本: {anchor_text!r}')


def _new_para_after(doc, after_idx: int) -> int:
    """在第 after_idx 段后插入空段落，返回新段落的 1-based 序号。"""
    rng = doc.Paragraphs(after_idx).Range
    rng.Collapse(_WD_COLLAPSE_END)
    rng.InsertParagraphAfter()
    return after_idx + 1


def _fmt(para, *, center: bool = True, indent: float = 0, keep_next: bool = False) -> None:
    """一次性设置段落对齐、首行缩进、左缩进、与下段同页。"""
    para.Alignment = _WD_ALIGN_CENTER if center else 0
    # 先清字符单位缩进（WPS 段落对话框"特殊格式→度量值"对应 firstLineChars）
    # 再清 points 单位缩进，两套都清才能保证"特殊格式=无"
    para.Format.CharacterUnitFirstLineIndent = 0
    para.Format.CharacterUnitLeftIndent = 0
    para.Format.FirstLineIndent = indent
    para.Format.LeftIndent = 0
    para.Format.RightIndent = 0
    para.Format.KeepWithNext = keep_next


def _add_pic(doc, para_idx: int, img_path: str, w_pt: float, h_pt: float,
             *, after_content: bool = False):
    """
    在第 para_idx 段中插入内嵌图片。

    after_content=False：插在段落开头（第一张图）。
    after_content=True ：插在段落现有内容末尾、段落标记之前（第二张图）。
    """
    if after_content:
        # End - 1 定位到段落标记前一个字符位置
        end = doc.Paragraphs(para_idx).Range.End - 1
        rng = doc.Range(end, end)
    else:
        rng = doc.Paragraphs(para_idx).Range
        rng.Collapse(_WD_COLLAPSE_START)

    shape = rng.InlineShapes.AddPicture(os.path.abspath(img_path), False, True)
    shape.Width = w_pt
    shape.Height = h_pt
    return shape


# ── 字符串处理（与 wps_ui 保持一致） ─────────────────────────────────────────────

def _normalize_captions(captions: list[str], chapter: int, fig_start: int) -> list[str]:
    """将 '(a) 内容' 格式图题升格为 '图 X-X  内容'，其余保留原样。"""
    result = []
    for i, cap in enumerate(captions):
        if re.match(r'^\([a-zA-Z]\)\s', cap):
            content = re.sub(r'^\([a-zA-Z]\)\s+', '', cap)
            result.append(f"图 {chapter}-{fig_start + i}  {content}")
        else:
            result.append(cap)
    return result


def _split_captions_to_lines(labels: list[str], captions: list[str]) -> list[str]:
    """将子图题按 MAX_LINE_CHARS 限制合并为若干行字符串。"""
    lines: list[str] = []
    cur: list[str] = []
    cur_len = 0
    for label, cap in zip(labels, captions):
        item = cap
        needed = len(item) + (1 if cur else 0)
        if cur and cur_len + needed > MAX_LINE_CHARS:
            lines.append(' '.join(cur))
            cur, cur_len = [item], len(item)
        else:
            cur.append(item)
            cur_len += needed
    if cur:
        lines.append(' '.join(cur))
    return lines


# ── 公开函数 ────────────────────────────────────────────────────────────────────

def insert_n_images_one_col(
    docx_path: str,
    anchor_text: str,
    images: list[str],
    captions: list[str],
    chapter: int,
    fig_start: int = 1,
    width: float = 12.0,
    height: float = 6.0,
    visible: bool = False,
) -> None:
    """
    单列图片插入：每张图独占一行，图题紧随其后，全部居中。

    Args:
        docx_path   : 目标文档路径
        anchor_text : 定位插入位置的段落文字
        images      : 图片路径列表
        captions    : 图题列表，与 images 等长
        chapter     : 章号（用于图题格式化，如 图 3-1）
        fig_start   : 本批图片的起始编号，默认 1
        width       : 图片宽度（厘米），默认 12.0
        height      : 图片高度（厘米），默认 6.0
        visible     : 是否显示 WPS 窗口（调试用）
    """
    abs_path = os.path.abspath(docx_path)
    w_pt, h_pt = _cm(width), _cm(height)
    captions = _normalize_captions(captions, chapter, fig_start)
    n = len(images)

    import pythoncom
    app = _get_app(visible)
    try:
        doc = app.Documents.Open(abs_path)
        after = _find_anchor_para(doc, anchor_text)
        print(f'[INFO] 锚定段落序号: {after}')

        for i in range(n):
            print(f'→ 插入图 {i + 1}/{n}: {os.path.basename(images[i])}')

            # 图片段：居中，首行缩进=0，与下段同页（图和图题不分页）
            img_idx = _new_para_after(doc, after)
            _add_pic(doc, img_idx, images[i], w_pt, h_pt)
            _fmt(doc.Paragraphs(img_idx), center=True, indent=0, keep_next=True)
            after = img_idx

            # 图题段：居中，首行缩进=0
            cap_idx = _new_para_after(doc, after)
            doc.Paragraphs(cap_idx).Range.InsertBefore(captions[i])
            _fmt(doc.Paragraphs(cap_idx), center=True, indent=0, keep_next=False)
            after = cap_idx

        doc.Save()
        doc.Close()
        print(f'[DONE] 已保存: {abs_path}')

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
        pythoncom.CoUninitialize()


def insert_n_images_two_col(
    docx_path: str,
    anchor_text: str,
    images: list[str],
    captions: list[str],
    total_caption: str,
    width: float = 6.99,
    height: float = 4.99,
    visible: bool = False,
) -> None:
    """
    双列图片插入：两张图横排在同一段落，上方标签行，下方子图题 + 总图题，全部居中。

    Args:
        docx_path     : 目标文档路径
        anchor_text   : 定位插入位置的段落文字
        images        : 图片路径列表（偶数长度为佳，奇数时最后一行只有一张）
        captions      : 子图题列表，与 images 等长
        total_caption : 总图题文字
        width         : 单张图片宽度（厘米），默认 6.99
        height        : 单张图片高度（厘米），默认 4.99
        visible       : 是否显示 WPS 窗口（调试用）
    """
    import pythoncom
    abs_path = os.path.abspath(docx_path)
    w_pt, h_pt = _cm(width), _cm(height)
    n = len(images)
    labels = [f"({chr(ord('a') + i)})" for i in range(n)]
    sub_lines = _split_captions_to_lines(labels, captions)

    app = _get_app(visible)
    try:
        doc = app.Documents.Open(abs_path)
        after = _find_anchor_para(doc, anchor_text)
        print(f'[INFO] 锚定段落序号: {after}')

        for i in range(0, n, 2):
            # 标签行：(a)   [33空格]   (b)，与下段同页
            label_text = labels[i]
            if i + 1 < n:
                label_text += ' ' * LABEL_GAP_SPACES + labels[i + 1]

            label_idx = _new_para_after(doc, after)
            doc.Paragraphs(label_idx).Range.InsertBefore(label_text)
            _fmt(doc.Paragraphs(label_idx), center=False, indent=0, keep_next=True)
            after = label_idx

            # 图片行：两张内嵌图片在同一段落
            img_idx = _new_para_after(doc, after)
            _add_pic(doc, img_idx, images[i], w_pt, h_pt, after_content=False)
            if i + 1 < n:
                _add_pic(doc, img_idx, images[i + 1], w_pt, h_pt, after_content=True)
            _fmt(doc.Paragraphs(img_idx), center=True, indent=0, keep_next=False)
            after = img_idx
            print(f'→ 插入图 {i + 1}' + (f', {i + 2}' if i + 1 < n else ''))

        # 子图题各行
        for sub_line in sub_lines:
            sub_idx = _new_para_after(doc, after)
            doc.Paragraphs(sub_idx).Range.InsertBefore(sub_line)
            _fmt(doc.Paragraphs(sub_idx), center=True, indent=0, keep_next=True)
            after = sub_idx

        # 总图题
        total_idx = _new_para_after(doc, after)
        doc.Paragraphs(total_idx).Range.InsertBefore(total_caption)
        _fmt(doc.Paragraphs(total_idx), center=True, indent=0, keep_next=False)

        doc.Save()
        doc.Close()
        print(f'[DONE] 已保存: {abs_path}')

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
