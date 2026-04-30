"""
Layer 3 — 单列图片排版工作流（ArUco 锚定图版）。

使用一张 ArUco 锚定图作为唯一 CV2 识别目标，
用户图片全程通过 Tab / ^g 键盘跳转操作，
避免因图片内容/颜色/尺寸差异导致识别失败。

输入：
  docx_path    : 文档路径
  anchor_text  : 定位段落文字
  anchor_image : ArUco 锚定图绝对路径（仅用于 CV2 定位，写死测试）
  images       : 用户图片路径列表
  captions     : 每张图的图题，与 images 等长
"""
from .. import wps_nav as W
from .. import primitives as P

HIT_DEFAULT_SINGLE_IMG_WIDTH  = 12.00
HIT_DEFAULT_SINGLE_IMG_HEIGHT = 6.00


def _find_text_backward(text: str) -> None:
    """Ctrl+F → 输入文本 → Alt+B（找上一处）→ Esc。
    用于光标已在锚定文本上时，避免向下搜索报找不到。"""
    P.hotkey('ctrl', 'f')
    P.wait(0.1)
    P.type_text(text)
    P.hotkey('alt', 'b')
    P.wait(0.1)
    P.press('escape')


def _find_next_graphic() -> None:
    """Ctrl+F → 输入 ^g → Enter → Esc，跳到文档中下一张内嵌图片。"""
    P.hotkey('ctrl', 'f')
    P.wait(0.1)
    P.type_text('^g')
    P.press('return')
    P.wait(0.5)
    P.press('escape')


def _set_center_align() -> None:
    """先左对齐再居中，确保任意初始格式都能正确居中。"""
    P.hotkey('alt')
    P.wait(0.1)
    P.hotkey('H')
    P.wait(0.1)
    P.hotkey('A')
    P.wait(0.1)
    P.hotkey('L')
    P.wait(0.1)
    P.hotkey('alt')
    P.wait(0.1)
    P.hotkey('H')
    P.wait(0.1)
    P.hotkey('A')
    P.wait(0.1)
    P.hotkey('C')
    P.wait(0.1)


def clear_indent() -> None:
    P.wait(0.1)
    P.hotkey('home')
    P.wait(0.1)
    P.hotkey('shift', 'end')
    P.wait(0.1)
    P.hotkey('alt')
    P.wait(0.1)
    P.hotkey('o')
    P.wait(0.1)
    P.hotkey('p')
    P.wait(0.1)
    P.hotkey('y')
    P.wait(0.1)
    P.hotkey('0')
    P.wait(0.1)
    P.hotkey('enter')
    P.wait(0.1)
    P.hotkey('end')


def _insert_image(path: str) -> None:
    W.open_insert_picture_dialog()
    W.input_file_path_confirm(path)
    P.wait(0.3)
    P.hotkey('right')  # 图片插入后处于选中态，right 移至图片后文本位置


def _set_keep_with_next() -> None:
    """为当前段落设置「与下段同页」，防止图片与图题跨页分离。"""
    P.hotkey('alt')
    P.wait(0.1)
    P.hotkey('o')
    P.wait(0.1)
    P.hotkey('p')
    P.wait(0.3)
    P.hotkey('p')
    P.wait(0.2)
    P.hotkey('x')
    P.wait(0.1)
    P.hotkey('i')
    P.wait(0.1)
    P.press('return')
    P.wait(0.2)


def _normalize_captions(captions: list[str], chapter: int, fig_start: int) -> list[str]:
    """将子图题格式 '(n) 内容' 升格为 '图 X-X  内容'，其余保留原样。"""
    import re
    result = []
    for i, cap in enumerate(captions):
        if re.match(r'^\([a-zA-Z]\)\s', cap):
            content = re.sub(r'^\([a-zA-Z]\)\s+', '', cap)
            result.append(f"图 {chapter}-{fig_start + i}  {content}")
        else:
            result.append(cap)
    return result


def _fig_prefix(caption: str) -> str:
    """从图题中提取 '图 X-X' 前缀用于 Ctrl+F 搜索。"""
    import re
    m = re.match(r'^(图\s+\d+-\d+)', caption)
    return m.group(1) if m else caption


def insert_n_images_one_col(
    docx_path: str,
    anchor_text: str,
    anchor_image: str,
    images: list[str],
    captions: list[str],
    chapter: int,
    fig_start: int = 1,
    debug: bool = False,
    run_phases: tuple[int, ...] = (1, 2, 3, 4, 5),
    width: float | None = None,
    height: float | None = None,
) -> None:
    from pathlib import Path
    from datetime import datetime

    try:
        n = len(images)
        captions = _normalize_captions(captions, chapter, fig_start)

        # ── Debug 初始化 ───────────────────────────────────────────────────────
        debug_dir = None
        _step = [0]
        if debug:
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            debug_dir = Path(__file__).parent.parent.parent.parent / 'debug' / ts
            debug_dir.mkdir(parents=True, exist_ok=True)
            print(f"[DEBUG] 截图目录：{debug_dir}")

        def snap(desc: str) -> None:
            if debug_dir is None:
                return
            _step[0] += 1
            path = debug_dir / f'{_step[0]:03d}_{desc}.png'
            P.screenshot().save(str(path))
            print(f'   [DEBUG] → {path.name}')

        # ── Phase 1: 插入所有内容 ──────────────────────────────────────────────
        if 1 in run_phases:
            W.open_doc(docx_path)
            W.goto_start()
            P.wait(1)
            snap('p1_doc_opened')

            print(f"→ 定位锚定文本：{anchor_text!r}")
            W.find_text(anchor_text)
            W.goto_line_end()
            snap('p1_anchor_found')

            print("→ 插入 ArUco 锚定图")
            P.hotkey('enter')
            _insert_image(anchor_image)
            snap('p1_anchor_image_inserted')

            print(f"→ 插入 {n} 张用户图片（单列布局）")
            for i in range(n):
                P.hotkey('enter')
                _set_keep_with_next()
                _insert_image(images[i])
                P.hotkey('enter')
                P.type_text(captions[i])
                clear_indent()
                snap(f'p1_image_{i + 1}')

            snap('p1_done')

        # ── Phase 2: 调整图片尺寸 ─────────────────────────────────────────────
        if 2 in run_phases:
            print("→ Ctrl+F 定位锚定文本，准备调整图片尺寸")
            _find_text_backward(anchor_text)
            P.wait(0.5)

            print("→ ^g 定位锚定图，CV2 点击，一次性导航到裁剪界面")
            _find_next_graphic()
            snap('p2_before_cv2_click')
            P.find_and_click_image(anchor_image, confidence=0.7)
            P.wait(0.3)
            import pyautogui
            pyautogui.rightClick()
            P.wait(0.3)
            P.press('o')
            W.navigate_to_crop_inputs(click_crop=True)
            snap('p2_property_panel_on_crop')

            _w = width if width is not None else HIT_DEFAULT_SINGLE_IMG_WIDTH
            _h = height if height is not None else HIT_DEFAULT_SINGLE_IMG_HEIGHT
            print(f"→ Tab 逐图填值，共 {n} 张（面板保持在裁剪界面，无需重复 CV2）")
            for i in range(n):
                print(f"   图 {i + 1}/{n}")
                P.hotkey('tab')
                P.wait(0.3)
                W.set_image_size_by_shortcut(
                    img_width=_w,
                    img_height=_h,
                    crop_width=_w,
                    crop_height=_h,
                )
                P.wait(0.5)
                if i == 0 or i == n - 1:
                    snap(f'p2_img_{i + 1}_resized')
            snap('p2_done')

        # ── Phase 3: 清除图片前多余字符，修复居中 ─────────────────────────────
        if 3 in run_phases:
            print("→ ^g 定位锚定图，跳过锚定图，准备清理图片前字符")
            _find_text_backward(anchor_text)
            _find_next_graphic()
            snap('p3_anchor_located')

            print(f"→ ^g 逐图清理，共 {n} 张")
            for i in range(n):
                print(f"   图 {i + 1}/{n}")
                _find_next_graphic()
                P.hotkey('left')
                P.hotkey('backspace')
                P.hotkey('home')
                P.hotkey('shift', 'end')
                _set_center_align()
                snap(f'p3_img_{i + 1}_centered')
                P.hotkey('right')
                P.hotkey('right')
            snap('p3_done')

        # ── Phase 4: 图题居中 ─────────────────────────────────────────────────
        if 4 in run_phases:
            print(f"→ 逐图题居中，共 {n} 行")
            for i, caption in enumerate(captions):
                W.find_text(_fig_prefix(caption))
                P.wait(0.3)
                P.hotkey('home')
                P.hotkey('shift', 'end')
                _set_center_align()
                snap(f'p4_caption_{i + 1}_centered')
            snap('p4_done')

        # ── Phase 5: 删除锚定图，保存关闭 ────────────────────────────────────
        if 5 in run_phases:
            print("→ 删除 ArUco 锚定图")
            _find_text_backward(anchor_text)
            _find_next_graphic()
            snap('p5_before_delete')
            P.press('delete')
            P.wait(0.3)
            snap('p5_anchor_deleted')

            print("→ 保存关闭")
            W.save_close()
            print("完成！")

    except Exception as e:
        snap('exception')
        print(e)
