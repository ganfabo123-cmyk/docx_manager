"""
image_solve.py — Layout-aware image insertion via COM (WPS / Word)

公开 API
────────
  ImageClass                  图片对象（图 + caption 封装）
  insert_image_after_para     插入图片到全文第 para_idx 段之后
  insert_image_within_section 在第 sec_idx 节内寻找合适位置插入图片

空间感知原理
────────────
  打开文档后 WPS/Word 完成排版计算；
  取第 para_idx+1 段起始 Y（含段后间距）作为第 para_idx 段的底边 Y；
  可用空间 = (PageHeight - BottomMargin) - para_bottom_Y；
  若 available < max_image_height + padding → 先插入分页符。

多图并排
────────
  传入 List[ImageClass] 时自动创建无边框表格；
  若各图宽之和 + 间距 > 版心宽，等比缩小所有图片；
  gap_pt 通过每列左右 padding 实现。
"""

from __future__ import annotations

import base64
import os
import re
import sys
import tempfile
from dataclasses import dataclass, field
from typing import List, Optional, Union

# ── COM 常量 ───────────────────────────────────────────────────────────────────
_WD_COLLAPSE_START    = 1
_WD_COLLAPSE_END      = 0
_WD_ALIGN_LEFT        = 0
_WD_ALIGN_CENTER      = 1
_WD_ALIGN_RIGHT       = 2
_WD_PRINT_VIEW        = 3
_WD_PAGE_BREAK        = 7        # wdPageBreak
_WD_VERT_POS_REL_PAGE = 6        # wdVerticalPositionRelativeToPage

_DEFAULT_PADDING_PT   = 18.0     # 图片与页面底边最小安全边距
_DEFAULT_GAP_PT       = 12.0     # 并排图片横向间距
_CAPTION_FONT_PT      = 10.5     # 图注字号（学术论文惯例）
_DEFAULT_FIG_REF_RE   = r'[如见]?图\s*\d+[-—–]\d+'


# ── ImageClass ─────────────────────────────────────────────────────────────────

@dataclass
class ImageClass:
    """
    图片对象，封装图片数据与 caption。

    image     : 图片文件路径 | 原始字节 | base64 字符串
    width_pt  : 显示宽度（points）
    height_pt : 显示高度（points）
    caption   : 图注文字（None = 不添加）
    """
    image:     Union[str, bytes]
    width_pt:  float
    height_pt: float
    caption:   Optional[str] = field(default=None)


# ── COM 应用管理 ───────────────────────────────────────────────────────────────

def _get_app(app_hint: str = 'auto', visible: bool = False):
    try:
        import win32com.client as wc
    except ImportError:
        print('[ERROR] pywin32 未安装: pip install pywin32', file=sys.stderr)
        sys.exit(1)

    order = {
        'wps':  ['Kwps.Application'],
        'word': ['Word.Application'],
        'auto': ['Kwps.Application', 'Word.Application'],
    }.get(app_hint, ['Kwps.Application', 'Word.Application'])

    for prog_id in order:
        try:
            app = wc.Dispatch(prog_id)
            app.Visible = visible
            print(f'[COM] 连接: {prog_id}')
            return app
        except Exception:
            continue

    print(f'[ERROR] 无法启动 COM 应用，已尝试: {order}', file=sys.stderr)
    sys.exit(1)


# ── 图片来源规范化 ─────────────────────────────────────────────────────────────

def _detect_suffix(data: bytes) -> str:
    if data[:4] == b'\x89PNG':                       return '.png'
    if data[:2] == b'\xff\xd8':                      return '.jpg'
    if data[:4] in (b'II\x2a\x00', b'MM\x00\x2a'):  return '.tiff'
    if data[:4] == b'\xd7\xcd\xc6\x9a':             return '.wmf'
    if data[:6] in (b'GIF87a', b'GIF89a'):           return '.gif'
    return '.png'


def _to_temp_file(image: Union[str, bytes]) -> tuple[str, bool]:
    """
    将 image 解析为磁盘文件路径。返回 (abs_path, is_temp)。

    修复点：优先判断文件路径是否存在，再尝试 base64 解码，
    避免"短 base64 字符串被误判为不存在的路径"的问题。
    """
    if isinstance(image, bytes):
        raw = image
    elif isinstance(image, str) and image.startswith('data:'):
        _, payload = image.split(',', 1)
        raw = base64.b64decode(payload)
    elif isinstance(image, str):
        # 优先：尝试文件路径
        if os.path.exists(image):
            return os.path.abspath(image), False
        # 回退：尝试 base64 解码
        try:
            raw = base64.b64decode(image, validate=True)
        except Exception:
            raise ValueError(f'image 既不是有效文件路径，也不是有效 base64: {image[:60]}')
    else:
        raise TypeError(f'不支持的 image 类型: {type(image)}')

    suffix = _detect_suffix(raw)
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    with open(path, 'wb') as f:
        f.write(raw)
    return path, True


# ── 空间感知 ───────────────────────────────────────────────────────────────────

def _ensure_print_view(app, doc) -> None:
    """
    切换到打印布局视图以激活版面坐标计算。

    修复点：visible=False 时 ActiveWindow 可能不可用，导致
    Information() 静默返回 0（空间感知永远不触发分页）。
    解决方案：临时将 app.Visible 设为 True 以访问窗口视图，操作后恢复。
    """
    try:
        was_visible = bool(app.Visible)
        if not was_visible:
            app.Visible = True
        doc.ActiveWindow.View.Type = _WD_PRINT_VIEW
        if not was_visible:
            app.Visible = False
    except Exception:
        pass


def _para_bottom_y(doc, global_idx: int) -> float:
    """
    返回第 global_idx 段底边 Y（points，从页顶量起）。
    取下一段起始 Y，自动包含段后间距。
    """
    total = doc.Paragraphs.Count
    if global_idx < total:
        rng = doc.Paragraphs(global_idx + 1).Range
        rng.Collapse(_WD_COLLAPSE_START)
    else:
        rng = doc.Paragraphs(global_idx).Range
        rng.Collapse(_WD_COLLAPSE_END)
    try:
        return float(rng.Information(_WD_VERT_POS_REL_PAGE))
    except Exception:
        return 0.0


def _available_space(doc, global_idx: int) -> float:
    """返回第 global_idx 段之后到页面内容区底边的可用垂直空间（points）。"""
    para = doc.Paragraphs(global_idx)
    try:
        ps = para.Range.Sections(1).PageSetup
    except Exception:
        ps = doc.Sections(1).PageSetup
    content_bottom = float(ps.PageHeight) - float(ps.BottomMargin)
    return content_bottom - _para_bottom_y(doc, global_idx)


def _get_content_width(doc, global_idx: int) -> float:
    """返回第 global_idx 段所在节的版心宽度（points）。"""
    try:
        ps = doc.Paragraphs(global_idx).Range.Sections(1).PageSetup
    except Exception:
        ps = doc.Sections(1).PageSetup
    return float(ps.PageWidth) - float(ps.LeftMargin) - float(ps.RightMargin)


# ── 段落全局索引查找 ───────────────────────────────────────────────────────────

def _global_idx_of(doc, para) -> int:
    """给定段落 COM 对象，返回其全文 1-based 全局序号。"""
    target = para.Range.Start
    total  = doc.Paragraphs.Count
    for i in range(1, total + 1):
        s = doc.Paragraphs(i).Range.Start
        if s == target:
            return i
        if s > target:
            return max(1, i - 1)
    return total


# ── 尺寸计算 ──────────────────────────────────────────────────────────────────

def _calc_scaled_sizes(
    img_list: List[ImageClass],
    content_width: float,
    gap_pt: float,
) -> tuple[List[float], List[float]]:
    """
    计算并排图片的缩放后宽高。
    若 sum(widths) + (n-1)*gap > content_width，等比缩小所有图片。
    """
    n            = len(img_list)
    orig_widths  = [img.width_pt  for img in img_list]
    orig_heights = [img.height_pt for img in img_list]
    total_gap    = (n - 1) * gap_pt
    total_img    = sum(orig_widths)
    available    = content_width - total_gap

    if total_img > available > 0:
        scale = available / total_img
        return [w * scale for w in orig_widths], [h * scale for h in orig_heights]
    return list(orig_widths), list(orig_heights)


# ── 分页插入 ───────────────────────────────────────────────────────────────────

def _insert_page_break_after(doc, para_idx: int) -> int:
    """在第 para_idx 段后插入分页符段落，返回分页符段落的全局索引。"""
    rng = doc.Paragraphs(para_idx).Range
    rng.Collapse(_WD_COLLAPSE_END)
    rng.InsertParagraphAfter()
    new_rng = doc.Paragraphs(para_idx + 1).Range
    new_rng.Collapse(_WD_COLLAPSE_START)
    new_rng.InsertBreak(_WD_PAGE_BREAK)
    return para_idx + 1


# ── 单图插入 ───────────────────────────────────────────────────────────────────

def _insert_single(
    doc,
    after_idx: int,
    img_obj:   ImageClass,
    width_pt:  float,
    height_pt: float,
    alignment: int,
) -> None:
    """在第 after_idx 段后插入单张图片，caption 紧随其后。"""
    rng = doc.Paragraphs(after_idx).Range
    rng.Collapse(_WD_COLLAPSE_END)
    rng.InsertParagraphAfter()
    img_para_idx = after_idx + 1

    img_rng = doc.Paragraphs(img_para_idx).Range
    img_rng.Collapse(_WD_COLLAPSE_START)

    img_path, is_temp = _to_temp_file(img_obj.image)
    try:
        shape        = img_rng.InlineShapes.AddPicture(img_path, False, True)
        shape.Width  = width_pt
        shape.Height = height_pt
    finally:
        if is_temp:
            try:
                os.unlink(img_path)
            except Exception:
                pass

    doc.Paragraphs(img_para_idx).Alignment = alignment

    if img_obj.caption:
        cap_rng = doc.Paragraphs(img_para_idx).Range
        cap_rng.Collapse(_WD_COLLAPSE_END)
        cap_rng.InsertParagraphAfter()
        cap_idx  = img_para_idx + 1
        cap_para = doc.Paragraphs(cap_idx)
        cap_para.Range.InsertBefore(img_obj.caption)
        cap_para.Alignment            = alignment
        cap_para.Range.Font.Size      = _CAPTION_FONT_PT

    print(f'[OK] 单图已插入（段 {img_para_idx}）'
          + (f'  caption: {img_obj.caption}' if img_obj.caption else ''))


# ── 多图并排插入（无边框表格） ────────────────────────────────────────────────

def _insert_table(
    doc,
    after_idx:     int,
    img_list:      List[ImageClass],
    scaled_widths: List[float],
    scaled_heights: List[float],
    gap_pt:        float,
    alignment:     int,
) -> None:
    """
    在第 after_idx 段后创建 1行N列无边框表格，每列放一张图。
    图注紧随图片写入同一单元格的第二段落。
    gap_pt 通过单元格左右 padding 实现（首列无左padding，末列无右padding）。
    """
    n = len(img_list)

    # 创建表格段落
    rng = doc.Paragraphs(after_idx).Range
    rng.Collapse(_WD_COLLAPSE_END)
    rng.InsertParagraphAfter()
    tbl_para_idx = after_idx + 1

    tbl_rng = doc.Paragraphs(tbl_para_idx).Range
    tbl_rng.Collapse(_WD_COLLAPSE_START)

    table = doc.Tables.Add(tbl_rng, 1, n)
    table.Borders.Enable = False
    try:
        table.Alignment = alignment
    except Exception:
        pass

    temp_files: List[str] = []

    for i, (img_obj, w, h) in enumerate(zip(img_list, scaled_widths, scaled_heights)):
        col_idx = i + 1
        try:
            table.Columns(col_idx).Width = w
        except Exception:
            pass

        cell = table.Cell(1, col_idx)

        # 设置列间距（左右 padding，首尾列单侧）
        try:
            cell.LeftPadding  = gap_pt / 2 if i > 0     else 0
            cell.RightPadding = gap_pt / 2 if i < n - 1 else 0
        except Exception:
            pass

        # 插入图片
        img_path, is_temp = _to_temp_file(img_obj.image)
        if is_temp:
            temp_files.append(img_path)

        cell_rng = cell.Range
        cell_rng.Collapse(_WD_COLLAPSE_START)
        shape        = cell_rng.InlineShapes.AddPicture(img_path, False, True)
        shape.Width  = w
        shape.Height = h

        cell.Range.Paragraphs(1).Alignment = alignment

        # caption 写入单元格第二段落
        if img_obj.caption:
            p1_rng = cell.Range.Paragraphs(1).Range
            p1_rng.Collapse(_WD_COLLAPSE_END)
            p1_rng.InsertParagraphAfter()
            cap_para = cell.Range.Paragraphs(2)
            cap_para.Range.InsertBefore(img_obj.caption)
            cap_para.Alignment       = alignment
            cap_para.Range.Font.Size = _CAPTION_FONT_PT

    # 清理临时文件
    for tmp in temp_files:
        try:
            os.unlink(tmp)
        except Exception:
            pass

    print(f'[OK] {n}图并排已插入（表格位于段 {tbl_para_idx}）')


# ── 统一调度 ───────────────────────────────────────────────────────────────────

def _do_insert(
    doc,
    after_global_idx: int,
    image:            Union[str, bytes, ImageClass, List[ImageClass]],
    width_pt:         Optional[float],
    height_pt:        Optional[float],
    caption:          Optional[str],
    gap_pt:           float,
    alignment:        int,
    padding_pt:       float,
) -> None:
    """
    统一入口：规范化图片列表 → 空间感知 → 分页或直接插入。
    """
    # ── 规范化为 List[ImageClass] ────────────────────────────────────────────
    if isinstance(image, list):
        img_list = image
    elif isinstance(image, ImageClass):
        img_list = [image]
    else:
        # str / bytes：需要 width_pt 和 height_pt
        if width_pt is None or height_pt is None:
            raise ValueError('image 为路径/字节时必须提供 width_pt 和 height_pt')
        img_list = [ImageClass(image, width_pt, height_pt, caption)]

    n = len(img_list)

    # ── 计算缩放尺寸 ─────────────────────────────────────────────────────────
    content_width                = _get_content_width(doc, after_global_idx)
    scaled_widths, scaled_heights = _calc_scaled_sizes(img_list, content_width, gap_pt)
    max_height                   = max(scaled_heights)

    # ── 空间感知 ─────────────────────────────────────────────────────────────
    available  = _available_space(doc, after_global_idx)
    need       = max_height + padding_pt
    insert_after = after_global_idx

    if available < need:
        print(f'[LAYOUT] 段 {after_global_idx} 后剩余 {available:.1f}pt '
              f'< 所需 {need:.1f}pt → 插入分页')
        insert_after = _insert_page_break_after(doc, after_global_idx)
    else:
        print(f'[LAYOUT] 段 {after_global_idx} 后剩余 {available:.1f}pt '
              f'≥ 所需 {need:.1f}pt → 直接插入')

    # ── 插入 ─────────────────────────────────────────────────────────────────
    if n == 1:
        _insert_single(doc, insert_after, img_list[0],
                       scaled_widths[0], scaled_heights[0], alignment)
    else:
        _insert_table(doc, insert_after, img_list,
                      scaled_widths, scaled_heights, gap_pt, alignment)


# ── 公开函数 ────────────────────────────────────────────────────────────────────

def insert_image_after_para(
    docx_path:  str,
    para_idx:   int,
    image:      Union[str, bytes, ImageClass, List[ImageClass]],
    width_pt:   Optional[float] = None,
    height_pt:  Optional[float] = None,
    caption:    Optional[str]   = None,
    padding_pt: float           = _DEFAULT_PADDING_PT,
    gap_pt:     float           = _DEFAULT_GAP_PT,
    alignment:  str             = 'center',
    app_hint:   str             = 'auto',
    visible:    bool            = False,
) -> None:
    """
    将图片插入到全文第 para_idx 段（1-based）之后。
    空间不足时自动插入分页符。

    image 可以是：
      str/bytes            单图路径或字节（需提供 width_pt / height_pt / caption）
      ImageClass           单图对象
      List[ImageClass]     多图并排
    """
    abs_path = os.path.abspath(docx_path)
    _align   = {'left': _WD_ALIGN_LEFT,
                'center': _WD_ALIGN_CENTER,
                'right': _WD_ALIGN_RIGHT}.get(alignment, _WD_ALIGN_CENTER)

    app = _get_app(app_hint, visible)
    try:
        doc   = app.Documents.Open(abs_path)
        _ensure_print_view(app, doc)

        total = doc.Paragraphs.Count
        if not (1 <= para_idx <= total):
            raise ValueError(f'para_idx={para_idx} 超出范围 [1, {total}]')

        _do_insert(doc, para_idx, image, width_pt, height_pt,
                   caption, gap_pt, _align, padding_pt)

        doc.Fields.Update()
        doc.Save()
        doc.Close()
        print(f'[DONE] 已保存: {abs_path}')

    except Exception as exc:
        print(f'[ERROR] {exc}', file=sys.stderr)
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


def insert_image_within_section(
    docx_path:       str,
    sec_idx:         int,
    image:           Union[str, bytes, ImageClass, List[ImageClass]],
    width_pt:        Optional[float] = None,
    height_pt:       Optional[float] = None,
    caption:         Optional[str]   = None,
    anchor:          str             = 'end',
    fig_ref_pattern: Optional[str]   = None,
    padding_pt:      float           = _DEFAULT_PADDING_PT,
    gap_pt:          float           = _DEFAULT_GAP_PT,
    alignment:       str             = 'center',
    app_hint:        str             = 'auto',
    visible:         bool            = False,
) -> None:
    """
    在第 sec_idx 节（1-based）内寻找合适段落，将图片插入其后。

    anchor='end'  → 节尾（兜底，不丢图）
    anchor='auto' → 扫描节内正文，找最后一个匹配 fig_ref_pattern 的段落；
                    无匹配时降级为 'end'
    """
    abs_path = os.path.abspath(docx_path)
    _align   = {'left': _WD_ALIGN_LEFT,
                'center': _WD_ALIGN_CENTER,
                'right': _WD_ALIGN_RIGHT}.get(alignment, _WD_ALIGN_CENTER)
    _re      = re.compile(fig_ref_pattern or _DEFAULT_FIG_REF_RE)

    app = _get_app(app_hint, visible)
    try:
        doc        = app.Documents.Open(abs_path)
        _ensure_print_view(app, doc)

        total_secs = doc.Sections.Count
        if not (1 <= sec_idx <= total_secs):
            raise ValueError(f'sec_idx={sec_idx} 超出范围 [1, {total_secs}]')

        sec       = doc.Sections(sec_idx)
        sec_paras = sec.Range.Paragraphs
        n_local   = sec_paras.Count

        if n_local == 0:
            raise ValueError(f'第 {sec_idx} 节没有可用段落')

        target_local = n_local  # 默认节尾

        if anchor == 'auto':
            last_match = None
            for i in range(1, n_local + 1):
                if _re.search(sec_paras(i).Range.Text):
                    last_match = i
            if last_match is not None:
                target_local = last_match
                print(f'[AUTO] 节 {sec_idx} 找到图文引用，锚定到局部段 {target_local}')
            else:
                print(f'[AUTO] 节 {sec_idx} 未找到图文引用，降级为节尾')

        global_idx = _global_idx_of(doc, sec_paras(target_local))
        print(f'[INFO] 目标段落: 节 {sec_idx} 局部 {target_local} → 全局 {global_idx}')

        _do_insert(doc, global_idx, image, width_pt, height_pt,
                   caption, gap_pt, _align, padding_pt)

        doc.Fields.Update()
        doc.Save()
        doc.Close()
        print(f'[DONE] 已保存: {abs_path}')

    except Exception as exc:
        print(f'[ERROR] {exc}', file=sys.stderr)
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
