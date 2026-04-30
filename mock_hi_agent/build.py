"""
完整流程编排

数据流一览
──────────────────────────────────────────────────────────────────
get_docs
    输入 : file_url
    输出 : job_id          → one_column / two_column_image / command_fixed_footer
           images          → ask_user_for_image_config (LLM节点)
           image_count     (informational)
           download_url    (无图无页码版，informational)

ask_user_for_image_config  [LLM节点，由 _call_llm_for_image_config 模拟]
    输入 : images           ← get_docs["images"]
                              每项: { index, anchor_text, caption }
           prompt           ← ask_user_for_image_config.PROMPT
    输出 : configured_images → one_column / two_column_image
           每项: { original_caption, display_caption, anchor_text, layout }

one_column
    输入 : job_id           ← get_docs["job_id"]
           configured_images ← _call_llm_for_image_config 输出 (仅 layout=="single")
           chapter          ← 外部参数
           fig_start        ← 外部参数
    输出 : download_url     (informational)
           inserted

two_column_image
    输入 : job_id           ← get_docs["job_id"]
           configured_images ← _call_llm_for_image_config 输出 (仅 layout=="two-col"，恰好 2 张)
           total_caption    ← 外部参数 (可选)
    输出 : download_url     (informational)
           inserted

command_fixed_footer
    输入 : job_id           ← get_docs["job_id"]
           body_section     ← 外部参数
    输出 : download_url     ← 最终文档下载链接
──────────────────────────────────────────────────────────────────
"""

import get_docs as _get_docs
import one_column as _one_column
import two_column_image as _two_col
import command_fixed_footer as _footer
import ask_user_for_image_config as _ask_cfg


def _call_llm_for_image_config(
    images: list[dict],   # 来自 get_docs["images"]，每项: { index, anchor_text, caption }
    prompt: str,          # ask_user_for_image_config.PROMPT
) -> list[dict]:
    """
    模拟 LLM 对话节点：将 images 和 prompt 发送给模型，
    收集用户对每张图的排版配置，返回 configured_images。

    真实实现应调用 LLM API，此处为 mock，直接按原始数据构造返回值。

    输入:
        images  : [{ "index": int, "anchor_text": str, "caption": str }, ...]
        prompt  : PROMPT 模板字符串（内含 {{ images }} 占位符）

    输出:
        configured_images : [
            {
                "original_caption": str,   # 不可修改，服务端查找 key（= images[i]["caption"]）
                "display_caption":  str,   # 写入文档的最终图题（可被用户改动）
                "anchor_text":      str,   # 定位段落文字（可被用户改动）
                "layout":           str,   # "single" 或 "two-col"
            },
            ...
        ]
    """
    # ── TODO: 替换为真实 LLM 调用 ───────────────────────────────
    # 示例: 将 prompt.replace("{{ images }}", json.dumps(images)) 发给模型
    # ─────────────────────────────────────────────────────────────

    # mock: 所有图默认 single，原样保留 caption 和 anchor_text
    return [
        {
            "original_caption": img["caption"],
            "display_caption":  img["caption"],
            "anchor_text":      img["anchor_text"],
            "layout":           "two-col",
        }
        for img in images
    ]


def build(
    file_url: str,
    chapter: int = 1,
    fig_start: int = 1,
    body_section: int = 4,
    total_caption: str = None,
) -> dict:
    # ── Step 1: get_docs ─────────────────────────────────────────
    # 输入 : file_url
    # 输出 : job_id / images / image_count / download_url
    docs = _get_docs.handler({"file_url": file_url})
    job_id      = docs["job_id"]        # 传递给后续所有节点
    images      = docs["images"]        # → Step 2
    image_count = docs["image_count"]

    # ── Step 2: ask_user_for_image_config (LLM节点) ──────────────
    # 输入 : images           ← docs["images"]
    #         prompt          ← ask_user_for_image_config.PROMPT
    # 输出 : configured_images → Step 3 / Step 4
    configured_images = _call_llm_for_image_config(
        images=images,
        prompt=_ask_cfg.PROMPT,
    )

    # ── Step 3: one_column ───────────────────────────────────────
    # 输入 : job_id           ← docs["job_id"]
    #         configured_images ← Step 2 输出 (handler 内部过滤 layout=="single")
    #         chapter / fig_start ← 外部参数
    # 输出 : download_url / inserted
    if False:
        one_col = _one_column.handler({
            "job_id":            job_id,
            "configured_images": configured_images,
            "chapter":           chapter,
            "fig_start":         fig_start,
        })

    # ── Step 4: two_column_image ─────────────────────────────────
    # 输入 : job_id           ← docs["job_id"]
    #         configured_images ← Step 2 输出 (handler 内部过滤 layout=="two-col"，恰好 2 张)
    #         total_caption   ← 外部参数
    # 输出 : download_url / inserted
    two_col = _two_col.handler({
        "job_id":            job_id,
        "configured_images": configured_images,
        "total_caption":     total_caption,
    })

    # ── Step 5: command_fixed_footer ─────────────────────────────
    # 输入 : job_id           ← docs["job_id"]
    #         body_section    ← 外部参数
    # 输出 : download_url     ← 最终可下载文档
    footer = _footer.handler({
        "job_id":       job_id,
        "body_section": body_section,
    })

    return {
        "job_id":       job_id,
        "image_count":  image_count,
        "images":       images,
      #  "one_col":      one_col,
        "two_col":      two_col,
        "download_url": footer["download_url"],
    }


if __name__ == "__main__":
    result = build(
        file_url="https://agent.hit.edu.cn/api/proxy/down?Action=Download&Version=2022-01-01&Path=upload%2Ffull%2F9f%2F42%2Fdb21d0b8515ed0db40eb2c3a3e3eb206720719c96955b1475f873ef45188&IsAnonymous=true",
        chapter=3,
        fig_start=1,
        body_section=4,
    )
    print(result)
