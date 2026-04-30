"""
Layer 3 — 双图横排排版工作流（ArUco 锚定图版）。

使用一张 ArUco 锚定图作为唯一 CV2 识别目标，
用户图片全程通过 Tab / ^g 键盘跳转操作，
避免因图片内容/颜色/尺寸差异导致识别失败。

输入：
  docx_path     : 文档路径
  anchor_text   : 定位段落文字
  anchor_image  : ArUco 锚定图绝对路径（仅用于 CV2 定位，写死测试）
  images        : 用户图片路径列表
  captions      : 每张图的子图题，与 images 等长
  total_caption : 总图题
"""
from .. import wps_nav as W
from .. import primitives as P

HIT_DEFAULT_DOUBLE_IMG_WIDTH  = 6.99
HIT_DEFAULT_DOUBLE_IMG_HEIGHT = 4.99
LABEL_GAP_SPACES = 33  # 标签行 (a) 与 (b) 之间的空格数
MAX_LINE_CHARS    = 32  # 子图题行最大字符数（按 len() 计，中文字符算 1）



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


def _open_property_panel() -> None:
    """当前图片已被 Tab 选中时，通过键盘打开属性面板：Shift+F10 → O"""
    import pyautogui
    P.wait(2)
    pyautogui.hotkey('shift', 'f10')
    P.wait(2)
    P.press('o')
    P.wait(0.5)


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
    P.hotkey('shift','end')
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
    print("press right for the anchor image")
    P.hotkey('right')  # 图片插入后处于选中态，right 移至图片后文本位置


def _split_captions_to_lines(labels: list, captions: list) -> list[str]:
    """将子图题按 MAX_LINE_CHARS 限制拆分为若干行字符串。"""
    lines: list[str] = []
    current_items: list[str] = []
    current_len = 0
    for label, cap in zip(labels, captions):
        if False and not cap.startswith(label):
            item = f"{label} {cap}"
        else:
            item = cap
        needed = len(item) + (1 if current_items else 0)  # +1 for space between items
        if current_items and current_len + needed > MAX_LINE_CHARS:
            lines.append(' '.join(current_items))
            current_items = [item]
            current_len = len(item)
        else:
            current_items.append(item)
            current_len += needed
    if current_items:
        lines.append(' '.join(current_items))
    return lines


def _set_keep_with_next() -> None:
    """为当前段落设置「与下段同页」，防止标签行与图片行跨页分离。"""
    P.hotkey('alt')
    P.wait(0.1)
    P.hotkey('o')
    P.wait(0.1)
    P.hotkey('p')
    P.wait(0.3)   # 等段落对话框打开
    P.hotkey('p')  # 切换到「换行与分页」tab
    P.wait(0.2)
    P.hotkey('x')  # 勾选「与下段同页」
    P.wait(0.1)
    P.hotkey('i')
    P.wait(0.1)
    P.press('return')
    P.wait(0.2)


def insert_n_images_two_col(
    docx_path: str,
    anchor_text: str,
    anchor_image: str,
    images: list[str],
    captions: list[str],
    total_caption: str,
    debug: bool = False,
    run_phases: tuple[int, ...] = (1, 2, 3, 4, 5),
) -> None:
    from pathlib import Path
    from datetime import datetime

    try:
        n = len(images)
        num_rows = (n + 1) // 2
        labels = [f"({chr(ord('a') + i)})" for i in range(n)]
        sub_lines = _split_captions_to_lines(labels, captions)
        num_sub_lines = len(sub_lines)

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

            print(f"→ 插入 {n} 张用户图片（两列布局）")
            for i in range(0, n, 2):
                P.hotkey('enter')
                P.type_text(labels[i])
                if i + 1 < n:
                    P.type_text(' ' * LABEL_GAP_SPACES)
                    P.type_text(labels[i + 1])
                clear_indent()
                _set_keep_with_next()
                snap(f'p1_label_row_{i // 2}')
                P.hotkey('enter')
                _insert_image(images[i])
                if i + 1 < n:
                    _insert_image(images[i + 1])
                snap(f'p1_image_row_{i // 2}')

            print(f"→ 插入子图题（{num_sub_lines} 行）和总图题")
            for sub_line in sub_lines:
                P.hotkey('enter')
                P.type_text(sub_line)
                clear_indent()
            snap('p1_subcaptions_inserted')
            P.hotkey('enter')
            P.type_text(total_caption)
            clear_indent()
            P.wait(0.3)
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
            P.hotkey('shift', 'f10')
            P.wait(0.3)
            P.hotkey('o')
            # 只做 CV2 导航（panel_tab → image_btn → crop_btn），不填值
            W.navigate_to_crop_inputs(click_crop=True)
            snap('p2_property_panel_on_crop')

            print(f"→ Tab 逐图填值，共 {n} 张（面板保持在裁剪界面，无需重复 CV2）")
            for i in range(n):
                print(f"   图 {i + 1}/{n}")
                P.hotkey('tab')
                P.wait(0.3)
                W.set_image_size_by_shortcut(
                    img_width=HIT_DEFAULT_DOUBLE_IMG_WIDTH,
                    img_height=HIT_DEFAULT_DOUBLE_IMG_HEIGHT,
                    crop_width=HIT_DEFAULT_DOUBLE_IMG_WIDTH,
                    crop_height=HIT_DEFAULT_DOUBLE_IMG_HEIGHT,
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
                if i % 2 == 1 or i == n - 1:
                    P.hotkey('space')
                    P.hotkey('space')
                    P.hotkey('space')
                    P.hotkey('home')
                    P.hotkey('shift', 'end')
                    _set_center_align()
                    snap(f'p3_row_{i // 2}_centered')
                P.hotkey('right')
                P.hotkey('right')
            snap('p3_done')

        # ── Phase 4: 子图题与总图题居中 ──────────────────────────────────────
        if 4 in run_phases:
            print(f"→ 定位总图题：{total_caption!r}，调整 {num_sub_lines + 1} 行居中")
            W.find_text(total_caption)
            P.wait(0.3)
            snap('p4_total_caption_found')
            for _ in range(num_sub_lines):
                P.hotkey('up')
                P.wait(0.1)
            for _ in range(num_sub_lines):
                P.hotkey('home')
                P.hotkey('shift', 'end')
                _set_center_align()
                P.hotkey('down')
                P.wait(0.1)
            P.hotkey('home')
            P.hotkey('shift', 'end')
            _set_center_align()
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