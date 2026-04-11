"""
模拟客户端程序
模拟一个样式可能混乱的 docx 文档变成样式正确的文档的完整流程
"""
import os
import requests
import traceback
from typing import List, Dict, Any
from pydantic import BaseModel, Field
from openai import OpenAI
import instructor

LLM_API_KEY = "sk-54f3b9e43a1b44ed8eaaf8666da594ca"
LLM_API_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
LLM_MODEL = "qwen3-30b-a3b-instruct-2507"

SERVER_URL = "http://192.168.47.118:5000"

FULL_DOCX_FILE_URL = "https://agent.hit.edu.cn/api/proxy/down?Action=Download&Version=2022-01-01&Path=upload%2Ffull%2Fe8%2F40%2Fe2287dec7224ec7da65630ef712ab6c60650293ff2f5aa739aef02b7f45d&IsAnonymous=true"
PART_DOCX_FILE_URL = "https://agent.hit.edu.cn/api/proxy/down?Action=Download&Version=2022-01-01&Path=upload%2Ffull%2F75%2F41%2F6e569b142cb40f6edc625af45972af4aac0c93952f273b9055262029db8e&IsAnonymous=true"
TXT_FILE_URL = "https://agent.hit.edu.cn/api/proxy/down?Action=Download&Version=2022-01-01&Path=upload%2Ffull%2Fa8%2F8e%2F83d4b9ac8b1d4fbf0c17e5ece81250531f2d97616e80720c7853acd893e9&IsAnonymous=true"
SAMPLE_TEXT = """# 第一章 绪论

## 1.1 研究背景

随着人工智能技术的快速发展，**自然语言处理**已经成为计算机科学领域的重要研究方向。

### 1.1.1 问题提出

在实际应用中，我们面临着诸多挑战：

- 数据质量问题
- 模型泛化能力
- 计算资源限制

## 1.2 研究意义

本研究具有重要的理论价值和实践意义。

### 1.2.1 理论意义

本研究丰富了相关领域的理论基础。

### 1.2.2 实践意义

研究成果可广泛应用于实际生产环境。

# 第二章 文献综述

## 2.1 国内外研究现状

近年来，国内外学者在该领域进行了大量研究。

## 2.2 现有研究的不足

尽管已有大量研究，但仍存在一些不足之处。
"""


class StyleBlock(BaseModel):
    id: str = Field(description="Element ID, e.g., elem_1, elem_2")
    type: str = Field(description="Style type: heading1, heading2, heading3, or normal")
    content: str = Field(description="Text content")


class StyleAnalysis(BaseModel):
    styles: List[StyleBlock]


def create_llm_client():
    client = OpenAI(
        api_key=LLM_API_KEY,
        base_url=LLM_API_BASE_URL
    )
    return instructor.patch(client)


def parse_docx_file(url: str) -> Dict[str, Any]:
    print(f"[1/5] 解析 DOCX 文件: {url[:50]}...")
    response = requests.post(
        f"{SERVER_URL}/parse-docx-file",
        json={"url": url},
        proxies={"http": None, "https": None}
    )
    
    if response.status_code == 200:
        data = response.json()
        print(f"      解析成功: {data.get('text_count')} 个文本元素, {data.get('total_count')} 个总元素")
        return data
    else:
        print(f"      解析失败: {response.json()}")
        return {}


def get_short_blocks() -> tuple:
    print("[2/5] 获取短文本块...")
    response = requests.get(f"{SERVER_URL}/identify-short-blocks")
    
    if response.status_code == 200:
        data = response.json()
        short_blocks = data.get('short_blocks', [])
        total_blocks = data.get('total_blocks', {})
        print(f"      获取成功: {len(short_blocks)} 个短文本块")
        return short_blocks, total_blocks
    else:
        print(f"      获取失败: {response.json()}")
        return [], {}


def analyze_styles_with_llm(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    print(f"[3/5] LLM 分析样式 ({len(blocks)} 个块)...")
    
    client = create_llm_client()
    
    prompt = f"""分析以下文本块，判断每个文本块的样式类型：

{blocks}

请识别：
- heading1: 一级标题（如"第1章 绪论"、"摘要"、"目录"等章节标题）
- heading2: 二级标题（如"1.1 课题背景"、"4.1 引言"等）
- heading3: 三级标题（如"1.2.1 气体润滑轴承的发展"等）
- normal: 普通正文

返回每个块的 id、type 和 content。"""
    
    try:
        result = client.chat.completions.create(
            model=LLM_MODEL,
            response_model=StyleAnalysis,
            messages=[
                {"role": "system", "content": "你是一个文档样式分析专家，负责识别文本块的样式类型。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1
        )
        
        styles = [{"id": s.id, "type": s.type, "content": s.content} for s in result.styles]
        print(f"      分析完成: {len(styles)} 个样式")
        return styles
    except Exception as e:
        print(f"      LLM 分析失败: {e}")
        traceback.print_exc()
        return []


def backfill_styles(edited_elements: List[Dict[str, Any]]) -> Dict[str, Any]:
    print(f"[4/5] 回填样式 ({len(edited_elements)} 个元素)...")
    response = requests.post(
        f"{SERVER_URL}/backfill-styles",
        json={"edited_elements": edited_elements}
    )
    
    if response.status_code == 200:
        data = response.json()
        print(f"      回填成功")
        return data.get('data', {})
    else:
        print(f"      回填失败: {response.json()}")
        return {}


def restore_document(data: List[Dict[str, Any]] = None) -> str:
    print("[5/5] 还原 DOCX 文档...")
    response = requests.post(
        f"{SERVER_URL}/restore-document",
        json={"data": data} if data else {}
    )
    
    if response.status_code == 200:
        result = response.json()
        file_path = result.get('file_path', '')
        print(f"      还原成功: {file_path}")
        return file_path
    else:
        print(f"      还原失败: {response.json()}")
        return ""


def main():
    print("=" * 60)
    print("文档样式修复流程")
    print("=" * 60)
    
    try:
        result =  parse_docx_file(FULL_DOCX_FILE_URL)
        if not result:
            print("解析失败，退出...")
            return
        
        short_blocks, total_blocks = get_short_blocks()
        if not short_blocks:
            print("未找到短文本块，退出...")
            return
        
        styles = analyze_styles_with_llm(short_blocks)
        if not styles:
            print("样式分析失败，退出...")
            return
        
        updated_data = backfill_styles(styles)
        if not updated_data:
            print("样式回填失败，退出...")
            return
        
        file_path = restore_document(updated_data)
        if not file_path:
            print("文档还原失败，退出...")
            return
        
        print("\n" + "=" * 60)
        print("文档样式修复完成!")
        print(f"输出文件: {file_path}")
        print("=" * 60)
        
    except Exception as e:
        print(f"流程执行失败: {e}")
        traceback.print_exc()


def parse_text(text: str, remove_markdown: bool = True) -> Dict[str, Any]:
    print(f"[1/3] 解析文本 ({len(text)} 字符)...")
    response = requests.post(
        f"{SERVER_URL}/parse-text",
        json={"text": text, "remove_markdown": remove_markdown}
    )
    
    if response.status_code == 200:
        data = response.json()
        print(f"      解析成功: {data.get('text_count')} 个文本元素")
        return data
    else:
        print(f"      解析失败: {response.json()}")
        return {}


def analyze_text_styles(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    print(f"[2/3] LLM 分析样式 ({len(blocks)} 个块)...")
    
    client = create_llm_client()
    
    prompt = f"""分析以下文本块，判断每个文本块的样式类型：

{blocks}

请识别：
- heading1: 一级标题（如"第一章 绪论"、"摘要"、"目录"等章节标题）
- heading2: 二级标题（如"1.1 课题背景"、"4.1 引言"等）
- heading3: 三级标题（如"1.2.1 气体润滑轴承的发展"等）
- normal: 普通正文

返回每个块的 id、type 和 content。"""
    
    try:
        result = client.chat.completions.create(
            model=LLM_MODEL,
            response_model=StyleAnalysis,
            messages=[
                {"role": "system", "content": "你是一个文档样式分析专家，负责识别文本块的样式类型。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1
        )
        
        styles = [{"id": s.id, "type": s.type, "content": s.content} for s in result.styles]
        print(f"      分析完成: {len(styles)} 个样式")
        return styles
    except Exception as e:
        print(f"      LLM 分析失败: {e}")
        traceback.print_exc()
        return []


def text_to_docx(text: str, styles: List[Dict[str, Any]] = None, remove_markdown: bool = True) -> str:
    print("[3/3] 生成 DOCX 文档...")
    response = requests.post(
        f"{SERVER_URL}/text-to-docx",
        json={"text": text, "remove_markdown": remove_markdown, "styles": styles or []}
    )
    
    if response.status_code == 200:
        result = response.json()
        file_path = result.get('file_path', '')
        print(f"      生成成功: {file_path}")
        return file_path
    else:
        print(f"      生成失败: {response.json()}")
        return ""


def main_text_to_docx():
    print("=" * 60)
    print("文本转文档流程")
    print("=" * 60)
    
    try:
        result = parse_text(SAMPLE_TEXT, remove_markdown=True)
        if not result:
            print("解析失败，退出...")
            return
        
        response = requests.get(f"{SERVER_URL}/identify-short-blocks")
        if response.status_code == 200:
            data = response.json()
            short_blocks = data.get('short_blocks', [])
        else:
            short_blocks = []
        
        if short_blocks:
            styles = analyze_text_styles(short_blocks)
        else:
            print("      未找到短文本块，直接生成文档...")
            styles = []
        
        file_path = text_to_docx(SAMPLE_TEXT, styles, remove_markdown=True)
        if not file_path:
            print("文档生成失败，退出...")
            return
        
        print("\n" + "=" * 60)
        print("文本转文档完成!")
        print(f"输出文件: {file_path}")
        print("=" * 60)
        
    except Exception as e:
        print(f"流程执行失败: {e}")
        traceback.print_exc()


def parse_txt_from_url(url: str, remove_markdown: bool = True) -> Dict[str, Any]:
    print(f"[1/3] 解析 TXT 文件: {url[:50]}...")
    response = requests.post(
        f"{SERVER_URL}/parse-txt-file",
        json={"url": url, "remove_markdown": remove_markdown}
    )
    
    if response.status_code == 200:
        data = response.json()
        print(f"      解析成功: {data.get('text_count')} 个文本元素")
        return data
    else:
        print(f"      解析失败: {response.json()}")
        return {}


def main_txt_to_docx():
    print("=" * 60)
    print("TXT文件转文档流程")
    print("=" * 60)
    
    SAMPLE_TXT_URL = "https://example.com/sample.txt"
    
    try:
        result = parse_txt_from_url(TXT_FILE_URL, remove_markdown=True)
        if not result:
            print("解析失败，退出...")
            return
        
        response = requests.get(f"{SERVER_URL}/identify-short-blocks")
        if response.status_code == 200:
            data = response.json()
            short_blocks = data.get('short_blocks', [])
        else:
            short_blocks = []
        
        if short_blocks:
            styles = analyze_text_styles(short_blocks)
        else:
            print("      未找到短文本块，直接生成文档...")
            styles = []
        
        response = requests.get(f"{SERVER_URL}/identify-short-blocks")
        if response.status_code == 200:
            data = response.json()
            total_blocks = data.get('total_blocks', {})
            text_elements = total_blocks.get('text_elements', [])
        else:
            text_elements = []
        
        if text_elements:
            text_content = '\n'.join([e.get('content', '') for e in text_elements])
            file_path = text_to_docx(text_content, styles, remove_markdown=False)
            if not file_path:
                print("文档生成失败，退出...")
                return
            
            print("\n" + "=" * 60)
            print("TXT文件转文档完成!")
            print(f"输出文件: {file_path}")
            print("=" * 60)
        else:
            print("未获取到文本内容，退出...")
        
    except Exception as e:
        print(f"流程执行失败: {e}")
        traceback.print_exc()


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg == '--text':
            main_text_to_docx()
        elif arg == '--txt':
            main_txt_to_docx()
        else:
            print(f"未知参数: {arg}")
            print("用法:")
            print("  python mock_client.py          # DOCX文档样式修复流程")
            print("  python mock_client.py --text   # 文本转文档流程")
            print("  python mock_client.py --txt    # TXT文件转文档流程")
    else:
        main()
