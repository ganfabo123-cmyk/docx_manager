"""
Layer 3 — HIT 学位论文页码工作流。

规则（来自 hit_config.json hit_footer_rule）：
  节 1 ~ (body_section-1)  → 大写罗马数字
  节 body_section ~ 末尾   → dash-arabic（-1-, -2-, ...）
"""
from .. import wps_nav as W


def set_roman_upper_from_here() -> None:
    """从当前节开始：大写罗马，应用本节及以后"""
    W.open_page_number_dialog()
    W.page_dialog_move(up=10)           # 先复位：10次上确保回到阿拉伯数字初始状态
    W.page_dialog_move(down=3)          # 3次下 = 大写罗马
    W.page_dialog_apply_this_section_onward()
    W.confirm()


def set_dash_arabic_from_here() -> None:
    """从当前节开始：dash-arabic，应用本节及以后"""
    W.open_page_number_dialog()
    W.page_dialog_move(up=10)           # 先复位：10次上确保回到阿拉伯数字初始状态
    W.page_dialog_move(down=1)          # 1次下 = dash-arabic（以阿拉伯数字为基准）
    W.page_dialog_apply_this_section_onward()
    W.confirm()


def apply_hit_page_numbers(docx_path: str, body_section: int = 4, close_delay: float = 2.0) -> None:
    """
    完整工作流：
      1. 打开文档，跳到开头
      2. 全文设大写罗马
      3. 跳到正文节（绪论）
      4. 正文节及以后覆盖为 dash-arabic
      5. 保存，延迟 close_delay 秒后关闭
    """
    W.open_doc(docx_path)
    W.goto_start()

    print("→ 全文设大写罗马")
    set_roman_upper_from_here()

    jumps = body_section - 1
    print(f"→ 跳转到正文节（跳 {jumps} 次）")
    for i in range(jumps):
        W.jump_next_section()
        print(f"  跳第 {i + 1} 次")

    print("→ 正文节起设 dash-arabic")
    set_dash_arabic_from_here()

    print(f"→ 保存，{close_delay}s 后关闭")
    W.save_close(close_delay)
    print("完成！")



def apply_page_numbers(docx_path: str, close_delay: float = 2.0) -> None:
    """
    完整工作流：
      1. 打开文档，跳到开头
      2. 全文设大写罗马
      3. 跳到正文节（绪论）
      4. 正文节及以后覆盖为 dash-arabic
      5. 保存，延迟 close_delay 秒后关闭
    """
    W.open_doc(docx_path)
    W.goto_start()

    print("→ 全文起设 dash-arabic")
    set_dash_arabic_from_here()

    print(f"→ 保存，{close_delay}s 后关闭")
    W.save_close(close_delay)
    print("完成！")
