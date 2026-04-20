"""
docx_cli.py — DOCX 后处理 CLI（通过 COM 自动化驱动 WPS / Word 引擎）

用法
────
  python docx_cli.py <子命令> <docx路径> [参数...]

子命令
──────
  page-number   对指定节范围设置页码样式、位置与起始编号
  insert-image  插入图片到指定段落或节内（支持多图并排）

page-number 完整语法
────────────────────
  python docx_cli.py page-number <docx>
      <style>     arabic | roman | dash-arabic
      <position>  top-left | top-center | top-right |
                  bottom-left | bottom-center | bottom-right
      <sec_from>  生效起始节（1-based）
      <sec_to>    生效末尾节，含此节（1-based；0 = 最后一节）
      [start_idx] 起始节的起始页码（默认 1）
      [--app      wps | word | auto]   COM 应用选择，默认 auto（先尝试 WPS）
      [--visible]  显示应用窗口（调试用）
      [--no-unlink] 不断开与前一节的链接（慎用）

示例
────
  # 摘要~目录（第1~2节）用小写罗马数字，居中置于底部，从 i 开始
  python docx_cli.py page-number output.docx roman bottom-center 1 2 1

  # 正文（第3节起到最后）用 -n- 格式，居中置于底部，从 1 开始
  python docx_cli.py page-number output.docx dash-arabic bottom-center 3 0 1

insert-image 完整语法
────────────────────
  python docx_cli.py insert-image after-para <docx>
      <para_idx>  目标段落全局序号（1-based）
      <image>     图片路径（可多个，实现多图并排）
      <width_pt>  图片显示宽度（points）
      <height_pt> 图片显示高度（points）
      [--caption  TEXT]           单图注释
      [--captions TEXT]           多图注释，用 "||" 分隔，如 "(a) 正面||(b) 侧面"
      [--padding  N]              空间判断安全边距（默认 18pt）
      [--gap      N]              多图横向间距（默认 12pt）
      [--align    left|center|right]
      [--app      wps|word|auto]
      [--visible]

  python docx_cli.py insert-image within-section <docx>
      <sec_idx>   目标节序号（1-based）
      <image>     图片路径（可多个）
      <width_pt>  图片显示宽度（points）
      <height_pt> 图片显示高度（points）
      [--caption  TEXT]           单图注释
      [--captions TEXT]           多图注释，用 "||" 分隔，如 "(a) 正面||(b) 侧面"
      [--anchor   end|auto]       锚定模式（end=节尾，auto=自动查找图文引用）
      [--fig-pattern REGEX]       图文引用正则（anchor=auto 时生效）
      [--padding  N]              空间判断安全边距（默认 18pt）
      [--gap      N]              多图横向间距（默认 12pt）
      [--align    left|center|right]
      [--app      wps|word|auto]
      [--visible]

insert-image 示例
─────────────────
  # 单图：在第5段后插入一张图
  python docx_cli.py insert-image after-para output.docx 5 img.png 200 150 --caption "图1"

  # 多图并排：在第5段后插入三张图（无注释）
  python docx_cli.py insert-image after-para output.docx 5 img1.png img2.png img3.png 200 150 --gap 12

  # 多图并排：在第5段后插入三张图（带各自注释）
  python docx_cli.py insert-image after-para output.docx 5 img1.png img2.png img3.png 200 150 --captions "(a) 正面||(b) 侧面||(c) 背面" --gap 12

  # 在第3节内自动查找图文引用位置插入图片
  python docx_cli.py insert-image within-section output.docx 3 img.png 200 150 --anchor auto
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import List, Optional

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

try:
    from ..utils.image_solve import (
        ImageClass,
        insert_image_after_para,
        insert_image_within_section,
        _DEFAULT_PADDING_PT,
        _DEFAULT_GAP_PT,
    )
except ImportError:
    from docx_manager.docx_engine.utils.image_solve import (
        ImageClass,
        insert_image_after_para,
        insert_image_within_section,
        _DEFAULT_PADDING_PT,
        _DEFAULT_GAP_PT,
    )

# ── COM 常量（不依赖 win32com.client.constants，保持可移植） ───────────────────
_WD_HEADER_FOOTER_PRIMARY = 1
_WD_HEADER_FOOTER_EVEN    = 3   # 奇偶页不同时的偶数页 header/footer（wdHeaderFooterEvenPages）
_WD_ALIGN_LEFT            = 0
_WD_ALIGN_CENTER          = 1
_WD_ALIGN_RIGHT           = 2
_WD_FIELD_PAGE            = 33   # wdFieldPage
_WD_COLLAPSE_END          = 0    # wdCollapseEnd

# wdPageNumberStyle
_STYLE_ARABIC       = 0   # 1, 2, 3 …
_STYLE_ROMAN_UPPER  = 1   # I, II, III …
_STYLE_ROMAN_LOWER  = 2   # i, ii, iii …

# style 参数 → (native_style | None, is_dash)
_STYLE_MAP = {
    'arabic':      (_STYLE_ARABIC,      False),
    'roman':       (_STYLE_ROMAN_LOWER, False),   # 小写罗马，学术论文惯例
    'roman-upper': (_STYLE_ROMAN_UPPER, False),
    'dash-arabic': (None,               True),    # -1-, -2-, … 自定义构建
}

# position 参数 → (location, alignment)
_POSITION_MAP = {
    'top-left':      ('header', _WD_ALIGN_LEFT),
    'top-center':    ('header', _WD_ALIGN_CENTER),
    'top-right':     ('header', _WD_ALIGN_RIGHT),
    'bottom-left':   ('footer', _WD_ALIGN_LEFT),
    'bottom-center': ('footer', _WD_ALIGN_CENTER),
    'bottom-right':  ('footer', _WD_ALIGN_RIGHT),
}


# ── COM 应用管理 ───────────────────────────────────────────────────────────────

def _get_com_app(app_hint: str, visible: bool):
    """
    获取 WPS 或 Word 的 COM 应用对象。

    返回 (app_object, prog_id_string)。
    app_hint: 'wps' | 'word' | 'auto'
    """
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        print('[ERROR] pywin32 未安装。请运行: pip install pywin32', file=sys.stderr)
        sys.exit(1)

    import win32com.client as wc

    if app_hint == 'wps':
        candidates = ['Kwps.Application']
    elif app_hint == 'word':
        candidates = ['Word.Application']
    else:
        candidates = ['Kwps.Application', 'Word.Application']

    for prog_id in candidates:
        try:
            app = wc.Dispatch(prog_id)
            app.Visible = visible
            print(f'[COM] 已连接: {prog_id}')
            return app, prog_id
        except Exception:
            continue

    print(f'[ERROR] 未能启动 COM 应用。已尝试: {candidates}', file=sys.stderr)
    print('        请确认 WPS Office 或 Microsoft Word 已安装。', file=sys.stderr)
    sys.exit(1)


# ── 页码构建核心 ───────────────────────────────────────────────────────────────

def _get_hf(section, location: str, even: bool = False):
    """根据 location ('header'|'footer') 返回 primary 或 even header/footer 对象。"""
    idx = _WD_HEADER_FOOTER_EVEN if even else _WD_HEADER_FOOTER_PRIMARY
    if location == 'header':
        return section.Headers(idx)
    return section.Footers(idx)


def _clear_hf(hf) -> None:
    """清空 header/footer 内容，保留段落结构。"""
    try:
        hf.Range.Text = ''
    except Exception:
        pass


def _apply_native_page_number(hf, doc, native_style: int, alignment: int) -> None:
    """
    arabic / roman 格式：使用 PageNumbers.Add 插入原生页码域。
    """
    _clear_hf(hf)
    try:
        for pn in hf.PageNumbers:
            pn.Delete()
    except Exception:
        pass
    hf.PageNumbers.Add(
        PageNumberAlignment=alignment,
        FirstPage=True,
    )
    hf.PageNumbers.NumberStyle = native_style
    hf.Range.ParagraphFormat.Alignment = alignment


def _apply_dash_arabic(hf, doc, alignment: int) -> None:
    """
    dash-arabic 格式："- {PAGE} -"。

    策略：优先用 chr(19)/chr(21) 直接写入域分隔符（绕开 WPS Fields.Add 的各种 bug）；
    若失败则回退到 hf.Range.Fields.Add；再失败则写 '?'（占位）。
    restart 由 _patch_restart_in_xml 在 XML 层处理。
    """
    _clear_hf(hf)

    # 方案 A：直接插入域字符 Chr(19)...Chr(21)，不经过 Fields.Add API
    # Chr(19)=field_begin, Chr(20)=field_sep, Chr(21)=field_end
    inserted = False
    try:
        rng = hf.Range
        rng.InsertAfter('- ' + chr(19) + ' PAGE ' + chr(21) + ' -')
        inserted = True
        print('[INFO] dash-arabic: 使用 chr(19/21) 插入域字符成功')
    except Exception as exc_a:
        print(f'[WARN] chr(19/21) 方式失败: {exc_a}，尝试 Fields.Add', file=sys.stderr)

    if not inserted:
        # 方案 B：hf.Range.Fields.Add（定向到本节 footer，非 doc 级别）
        try:
            _clear_hf(hf)
            rng = hf.Range
            rng.InsertAfter('- ')
            end_rng = hf.Range
            end_rng.Collapse(_WD_COLLAPSE_END)
            hf.Range.Fields.Add(end_rng, _WD_FIELD_PAGE)
            end_rng2 = hf.Range
            end_rng2.Collapse(_WD_COLLAPSE_END)
            end_rng2.InsertAfter(' -')
            inserted = True
            print('[INFO] dash-arabic: hf.Range.Fields.Add 成功')
        except Exception as exc_b:
            print(f'[WARN] hf.Range.Fields.Add 失败: {exc_b}，回退到占位符', file=sys.stderr)

    if not inserted:
        # 方案 C：占位符，XML patch 后续覆盖
        _clear_hf(hf)
        hf.Range.InsertAfter('- ? -')

    hf.Range.ParagraphFormat.Alignment = alignment


# ── page-number 命令实现 ───────────────────────────────────────────────────────

def _clear_all_hf_for_location(doc, location: str) -> None:
    """
    彻底清空全文所有节的 header/footer（primary + even）。

    LinkToPrevious=True 只是"显示前节内容"，不会清除该节自身存储的旧内容；
    后续 unlink 时旧内容会重新出现。因此必须逐节 unlink 后各自清空。
    """
    total = doc.Sections.Count
    for even in (False, True):
        for i in range(1, total + 1):
            sec = doc.Sections(i)
            hf = _get_hf(sec, location, even=even)
            try:
                hf.LinkToPrevious = False
            except Exception:
                pass
            hf = _get_hf(sec, location, even=even)  # re-fetch after unlink
            _clear_hf(hf)
    print(f'[INFO] 已清空全文 {location}（共 {total} 节，primary+even）')


def _parse_rule_str(s: str) -> dict:
    """
    解析 --rule 参数字符串，格式：
        sec_from-sec_to:style:position[:start_idx]
    示例：
        "1-3:roman-upper:bottom-center:1"
        "4-0:dash-arabic:bottom-center"
    """
    parts = s.split(':')
    if len(parts) < 3:
        raise ValueError(f'rule 格式错误: {s!r}，应为 "from-to:style:position[:start]"')
    range_part, style, position = parts[0], parts[1], parts[2]
    start_idx = int(parts[3]) if len(parts) >= 4 else 1
    if '-' not in range_part:
        raise ValueError(f'range 格式错误: {range_part!r}，应为 "from-to"')
    sec_from_s, sec_to_s = range_part.split('-', 1)
    return {
        'sec_from':  int(sec_from_s),
        'sec_to':    int(sec_to_s),
        'style':     style.lower(),
        'position':  position.lower(),
        'start_idx': start_idx,
    }


def _patch_restart_in_xml(docx_path: str, rules: list) -> None:
    """
    在 docx zip 内直接修改 document.xml：
    对每条规则的 sec_from 节，给对应的 inline sectPr 加上 <w:pgNumType w:start="N"/>。
    这是绕开 WPS COM RestartNumberingAtSection 不可靠问题的最终手段。

    COM Section N 对应 document.xml 中第 N 个 inline sectPr（在 <w:pPr> 里的 <w:sectPr>）。
    """
    import zipfile, re, shutil

    _W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    with zipfile.ZipFile(docx_path, 'r') as z:
        names = z.namelist()
        all_files = {n: z.read(n) for n in names}

    raw = all_files['word/document.xml'].decode('utf-8')

    # 注册所有命名空间前缀，防止 ET 改写 xmlns 声明
    for prefix, uri in re.findall(r'xmlns:(\w+)="([^"]+)"', raw):
        try:
            from xml.etree import ElementTree as ET
            ET.register_namespace(prefix, uri)
        except Exception:
            pass

    from xml.etree import ElementTree as ET

    def _q(tag):
        return f'{{{_W}}}{tag}'

    root = ET.fromstring(raw.encode('utf-8'))
    body = root.find(_q('body'))

    # 收集所有 inline sectPr（在 p/pPr 内的）
    inline_sectPrs = []
    for child in list(body):
        if child.tag != _q('p'):
            continue
        pPr = child.find(_q('pPr'))
        if pPr is None:
            continue
        sp = pPr.find(_q('sectPr'))
        if sp is not None:
            inline_sectPrs.append(sp)

    patched = 0
    for rule in rules:
        sec_from  = int(rule['sec_from'])
        start_idx = int(rule.get('start_idx', 1))
        idx = sec_from - 1  # 0-based

        if 0 <= idx < len(inline_sectPrs):
            sp = inline_sectPrs[idx]
            for pn in list(sp.findall(_q('pgNumType'))):
                sp.remove(pn)
            pn_elem = ET.SubElement(sp, _q('pgNumType'))
            pn_elem.set(_q('start'), str(start_idx))
            print(f'[XML]  第 {sec_from} 节 sectPr → pgNumType start="{start_idx}"')
            patched += 1
        else:
            print(f'[WARN] 第 {sec_from} 节超出 inline sectPr 范围（共 {len(inline_sectPrs)} 个），跳过')

    if not patched:
        return

    new_xml = ET.tostring(root, encoding='unicode')
    all_files['word/document.xml'] = new_xml.encode('utf-8')

    tmp = docx_path + '._tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)

    shutil.move(tmp, docx_path)
    print(f'[XML]  pgNumType patch 完成（{patched} 条规则）')


def cmd_page_number_multi(
    docx_path: str,
    rules:     list,
    app_hint:  str  = 'auto',
    visible:   bool = False,
) -> None:
    """
    一次性全文页码设置（幂等）：
      1. 清空所有规则涉及 location 的全文页眉/页脚
      2. 依次应用每条规则
    rules 每条: {sec_from, sec_to, style, position, start_idx}
    sec_to=0 表示最后一节。
    """
    abs_path = os.path.abspath(docx_path)
    if not os.path.exists(abs_path):
        print(f'[ERROR] 文件不存在: {abs_path}', file=sys.stderr)
        sys.exit(1)

    parsed: list = []
    for r in rules:
        style    = r['style'].lower()
        position = r['position'].lower()
        if style not in _STYLE_MAP:
            print(f'[ERROR] 不支持的 style: {style!r}，可选: {list(_STYLE_MAP)}',
                  file=sys.stderr)
            sys.exit(1)
        if position not in _POSITION_MAP:
            print(f'[ERROR] 不支持的 position: {position!r}，可选: {list(_POSITION_MAP)}',
                  file=sys.stderr)
            sys.exit(1)
        native_style, is_dash = _STYLE_MAP[style]
        location, alignment   = _POSITION_MAP[position]
        parsed.append({
            'sec_from':     int(r['sec_from']),
            'sec_to':       int(r['sec_to']),
            'style':        style,
            'position':     position,
            'start_idx':    int(r.get('start_idx', 1)),
            'native_style': native_style,
            'is_dash':      is_dash,
            'location':     location,
            'alignment':    alignment,
        })

    app, _ = _get_com_app(app_hint, visible)

    try:
        doc = app.Documents.Open(abs_path)
        total = doc.Sections.Count
        print(f'[INFO] 文档共 {total} 节')

        # Step 1: 清空全文（按 location 分组，避免重复清空）
        for location in {r['location'] for r in parsed}:
            _clear_all_hf_for_location(doc, location)

        # Step 2: 依次应用规则
        for rule in parsed:
            sec_from  = max(1, rule['sec_from'])
            sec_to    = total if rule['sec_to'] == 0 else min(rule['sec_to'], total)
            location  = rule['location']
            alignment = rule['alignment']

            if sec_from > sec_to:
                print(f'[WARN] 规则跳过: sec_from({sec_from}) > sec_to({sec_to})')
                continue

            for sec_idx in range(sec_from, sec_to + 1):
                sec = doc.Sections(sec_idx)

                restart = (sec_idx == sec_from)

                # 同时写 primary（奇数页）和 even（偶数页）
                for even in (False, True):
                    hf = _get_hf(sec, location, even=even)
                    try:
                        hf.LinkToPrevious = False
                    except Exception:
                        pass
                    hf = _get_hf(sec, location, even=even)  # re-fetch after unlink

                    if rule['is_dash']:
                        # restart 由 doc.Save() 后的 XML patch 负责，不经 COM
                        _apply_dash_arabic(hf, doc, alignment)
                    else:
                        _apply_native_page_number(hf, doc, rule['native_style'], alignment)
                        try:
                            pns = hf.PageNumbers
                            if restart:
                                pns.RestartNumberingAtSection = True
                                pns.StartingNumber = rule['start_idx']
                            else:
                                pns.RestartNumberingAtSection = False
                        except Exception as exc:
                            print(f'[WARN] 设置页码编号失败(even={even}): {exc}',
                                  file=sys.stderr)

                print(f'[OK]   第 {sec_idx} 节 → style={rule["style"]}, '
                      f'position={rule["position"]}')

        try:
            doc.Fields.Update()
            print('[INFO] Fields.Update() 完成')
        except Exception as exc:
            print(f'[WARN] Fields.Update() 失败: {exc}', file=sys.stderr)

        doc.Save()
        doc.Close()
        print(f'[DONE] COM 操作已保存: {abs_path}')

        # XML patch：直接写 pgNumType 到对应 sectPr，绕开 WPS COM RestartNumberingAtSection 可靠性问题
        _patch_restart_in_xml(abs_path, parsed)
        print(f'[DONE] 已完成: {abs_path}')

    except Exception as exc:
        import traceback
        print(f'[ERROR] 处理失败: {exc}', file=sys.stderr)
        traceback.print_exc()
        try:
            doc.Close(False)
        except Exception:
            pass
        sys.exit(1)
    finally:
        try:
            app.Quit()
        except Exception:
            pass


# ── insert-image 命令实现 ───────────────────────────────────────────────────────

def _parse_captions(captions_str: Optional[str], n_images: int) -> List[Optional[str]]:
    """
    解析 caption 字符串为列表。
    
    格式：用 "||" 分隔多个 caption，如 "(a) 正面||(b) 侧面||(c) 背面"
    若 captions_str 为 None 或数量不足，返回 [None, ...]
    """
    if not captions_str:
        return [None] * n_images
    
    captions = captions_str.split('||')
    
    if len(captions) < n_images:
        captions.extend([None] * (n_images - len(captions)))
    elif len(captions) > n_images:
        captions = captions[:n_images]
    
    return captions


def cmd_insert_image(args) -> None:
    """处理 insert-image 子命令。"""
    if args.subcmd == 'after-para':
        if len(args.image) == 1:
            if args.width_pt is None or args.height_pt is None:
                print('[ERROR] 单图模式需要提供 width_pt 和 height_pt', file=sys.stderr)
                sys.exit(1)
            insert_image_after_para(
                docx_path  = args.docx,
                para_idx   = args.para_idx,
                image      = args.image[0],
                width_pt   = args.width_pt,
                height_pt  = args.height_pt,
                caption    = args.caption,
                padding_pt = args.padding,
                gap_pt     = args.gap,
                alignment  = args.align,
                app_hint   = args.app,
                visible    = args.visible,
            )
        else:
            if args.width_pt is None or args.height_pt is None:
                print('[ERROR] 多图模式需要提供 width_pt 和 height_pt 作为基准尺寸', file=sys.stderr)
                sys.exit(1)
            captions = _parse_captions(args.captions, len(args.image))
            images = [
                ImageClass(img, args.width_pt, args.height_pt, cap)
                for img, cap in zip(args.image, captions)
            ]
            insert_image_after_para(
                docx_path  = args.docx,
                para_idx   = args.para_idx,
                image      = images,
                padding_pt = args.padding,
                gap_pt     = args.gap,
                alignment  = args.align,
                app_hint   = args.app,
                visible    = args.visible,
            )
    elif args.subcmd == 'within-section':
        if len(args.image) == 1:
            if args.width_pt is None or args.height_pt is None:
                print('[ERROR] 单图模式需要提供 width_pt 和 height_pt', file=sys.stderr)
                sys.exit(1)
            insert_image_within_section(
                docx_path       = args.docx,
                sec_idx         = args.sec_idx,
                image           = args.image[0],
                width_pt        = args.width_pt,
                height_pt       = args.height_pt,
                caption         = args.caption,
                anchor          = args.anchor,
                fig_ref_pattern = args.fig_pattern,
                padding_pt      = args.padding,
                gap_pt          = args.gap,
                alignment       = args.align,
                app_hint        = args.app,
                visible         = args.visible,
            )
        else:
            if args.width_pt is None or args.height_pt is None:
                print('[ERROR] 多图模式需要提供 width_pt 和 height_pt 作为基准尺寸', file=sys.stderr)
                sys.exit(1)
            captions = _parse_captions(args.captions, len(args.image))
            images = [
                ImageClass(img, args.width_pt, args.height_pt, cap)
                for img, cap in zip(args.image, captions)
            ]
            insert_image_within_section(
                docx_path       = args.docx,
                sec_idx         = args.sec_idx,
                image           = images,
                anchor          = args.anchor,
                fig_ref_pattern = args.fig_pattern,
                padding_pt      = args.padding,
                gap_pt          = args.gap,
                alignment       = args.align,
                app_hint        = args.app,
                visible         = args.visible,
            )


# ── CLI 入口 ──────────────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog='docx_cli',
        description='DOCX 后处理工具（COM 驱动 WPS / Word 引擎）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest='command', required=True)

    # ── page-number ────────────────────────────────────────────────────────────
    pn = sub.add_parser(
        'page-number',
        help='一次性全文页码方案（幂等，先清空再应用）',
    )
    pn.add_argument('docx', help='.docx 文件路径')
    pn.add_argument(
        '--rule',
        dest='rules', action='append', default=[],
        metavar='FROM-TO:STYLE:POSITION[:START]',
        help=(
            '页码规则，可重复使用多次。格式: "from-to:style:position[:start_idx]"。'
            '示例: "1-3:roman-upper:bottom-center:1"  "4-0:dash-arabic:bottom-center:1"'
        ),
    )
    pn.add_argument(
        '--from-config',
        dest='config_path', default=None,
        metavar='CONFIG_JSON',
        help='从 JSON 配置文件读取 hit_footer_rule.rules（与 --rule 合并）',
    )
    pn.add_argument('--app',     choices=['wps', 'word', 'auto'], default='auto',
                    help='指定 COM 应用（默认 auto，先尝试 WPS）')
    pn.add_argument('--visible', action='store_true',
                    help='显示应用窗口（调试用）')

    # ── insert-image ────────────────────────────────────────────────────────────
    img = sub.add_parser(
        'insert-image',
        help='插入图片到指定段落或节内（支持多图并排）',
    )
    img_sub = img.add_subparsers(dest='subcmd', required=True)

    # ── insert-image after-para ────────────────────────────────────────────────
    ap = img_sub.add_parser('after-para', help='插入图片到指定全局段落之后')
    ap.add_argument('docx')
    ap.add_argument('para_idx',  type=int, help='目标段落全局序号（1-based）')
    ap.add_argument('image',     nargs='+', help='图片路径（可多个）')
    ap.add_argument('width_pt',  type=float, nargs='?', help='图片宽度（points）')
    ap.add_argument('height_pt', type=float, nargs='?', help='图片高度（points）')
    ap.add_argument('--caption',    default=None, help='单图注释')
    ap.add_argument('--captions',   default=None, 
                    help='多图注释，用 "||" 分隔，如 "(a) 正面||(b) 侧面"')
    ap.add_argument('--padding',    type=float, default=_DEFAULT_PADDING_PT,
                    help='空间判断安全边距（默认 18pt）')
    ap.add_argument('--gap',        type=float, default=_DEFAULT_GAP_PT,
                    help='多图横向间距（默认 12pt）')
    ap.add_argument('--align',      default='center',
                    choices=['left', 'center', 'right'])
    ap.add_argument('--app',        default='auto',
                    choices=['wps', 'word', 'auto'])
    ap.add_argument('--visible',    action='store_true')

    # ── insert-image within-section ─────────────────────────────────────────────
    ws = img_sub.add_parser('within-section', help='在指定节内寻找合适位置插入图片')
    ws.add_argument('docx')
    ws.add_argument('sec_idx',   type=int, help='目标节序号（1-based）')
    ws.add_argument('image',     nargs='+', help='图片路径（可多个）')
    ws.add_argument('width_pt',  type=float, nargs='?', help='图片宽度（points）')
    ws.add_argument('height_pt', type=float, nargs='?', help='图片高度（points）')
    ws.add_argument('--caption',     default=None, help='单图注释')
    ws.add_argument('--captions',    default=None,
                    help='多图注释，用 "||" 分隔，如 "(a) 正面||(b) 侧面"')
    ws.add_argument('--anchor',      default='end', choices=['end', 'auto'],
                    help='锚定模式：end=节尾，auto=自动查找图文引用')
    ws.add_argument('--fig-pattern', default=None,
                    help='图文引用正则（anchor=auto 时生效）')
    ws.add_argument('--padding',     type=float, default=_DEFAULT_PADDING_PT,
                    help='空间判断安全边距（默认 18pt）')
    ws.add_argument('--gap',         type=float, default=_DEFAULT_GAP_PT,
                    help='多图横向间距（默认 12pt）')
    ws.add_argument('--align',       default='center',
                    choices=['left', 'center', 'right'])
    ws.add_argument('--app',         default='auto',
                    choices=['wps', 'word', 'auto'])
    ws.add_argument('--visible',     action='store_true')

    return parser


def main(argv: Optional[list[str]] = None) -> None:
    parser = _build_parser()
    args   = parser.parse_args(argv)

    if args.command == 'page-number':
        import json as _json
        rules: list = []
        # 先从 --from-config 读取
        if args.config_path:
            cfg_path = os.path.abspath(args.config_path)
            if not os.path.exists(cfg_path):
                print(f'[ERROR] 配置文件不存在: {cfg_path}', file=sys.stderr)
                sys.exit(1)
            with open(cfg_path, encoding='utf-8') as _f:
                _cfg = _json.load(_f)
            footer_rule = _cfg.get('hit_footer_rule', {})
            rules.extend(footer_rule.get('rules', []))
        # 再追加 --rule 参数
        for rule_str in args.rules:
            try:
                rules.append(_parse_rule_str(rule_str))
            except ValueError as e:
                print(f'[ERROR] {e}', file=sys.stderr)
                sys.exit(1)
        if not rules:
            print('[ERROR] 至少需要一条规则（--rule 或 --from-config）', file=sys.stderr)
            sys.exit(1)
        cmd_page_number_multi(
            docx_path = args.docx,
            rules     = rules,
            app_hint  = args.app,
            visible   = args.visible,
        )
    elif args.command == 'insert-image':
        cmd_insert_image(args)


if __name__ == '__main__':
    main()
