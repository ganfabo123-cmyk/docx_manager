import os
import requests
import traceback
from typing import List, Dict, Any
from pydantic import BaseModel, Field
from openai import OpenAI
import instructor

# LLM Configuration
LLM_API_KEY = "sk-54f3b9e43a1b44ed8eaaf8666da594ca"
LLM_API_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
LLM_MODEL = "qwen3-30b-a3b-instruct-2507"

# Server Configuration
SERVER_URL = "http://localhost:5000"

# File URL
FILE_URL = "https://agent.hit.edu.cn/api/proxy/down?Action=Download&Version=2022-01-01&Path=upload%2Ffull%2F66%2F9b%2Fc70cedf0e0971155db2a4849dd4b4198757825bcb8280ff32e71b4cc64ff&IsAnonymous=true"


class StyleBlock(BaseModel):
    id: int
    type: str = Field(description="Style type: h1, h2, h3, toc1, toc2, toc3, or paragraph")
    content: str


class StyleAnalysis(BaseModel):
    styles: List[StyleBlock]


def create_llm_client():
    """
    Create OpenAI client with instructor for structured output
    """
    client = OpenAI(
        api_key=LLM_API_KEY,
        base_url=LLM_API_BASE_URL
    )
    return instructor.patch(client)


def parse_docx_file(url: str) -> bool:
    """
    Parse DOCX file from URL
    """
    try:
        print(f"Parsing DOCX file from: {url}")
        response = requests.post(
            f"{SERVER_URL}/parse-docx-file",
            json={"url": url}
        )
        
        if response.status_code == 200:
            print("File parsed successfully")
            return True
        else:
            print(f"Failed to parse file: {response.json()}")
            return False
    except Exception as e:
        print(f"Error parsing file: {e}")
        print(traceback.format_exc())
        return False


def get_short_blocks() -> List[Dict[str, Any]]:
    """
    Get short blocks from parsed file
    """
    try:
        print("Getting short blocks...")
        response = requests.get(f"{SERVER_URL}/identify-short-blocks")
        
        if response.status_code == 200:
            data = response.json()
            short_blocks = data.get('short_blocks', [])
            print(f"Retrieved {len(short_blocks)} short blocks")
            return short_blocks
        else:
            print(f"Failed to get short blocks: {response.json()}")
            return []
    except Exception as e:
        print(f"Error getting short blocks: {e}")
        print(traceback.format_exc())
        return []


def analyze_styles_with_llm(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Analyze styles using LLM with instructor
    """
    try:
        print(f"Analyzing {len(blocks)} blocks with LLM...")
        
        client = create_llm_client()
        
        # Create prompt for LLM
        prompt = f"""Analyze the following short text blocks and determine their style type:

{blocks}

Please identify which blocks are:
- Level 1 heading (h1)
- Level 2 heading (h2)
- Level 3 heading (h3)
- Table of contents item (toc1, toc2, toc3)
- Other (leave as paragraph)

Return a structured response with id, type, and content for each block."""
        
        # Use instructor for structured output
        result = client.chat.completions.create(
            model=LLM_MODEL,
            response_model=StyleAnalysis,
            messages=[
                {"role": "system", "content": "You are a document style analyzer. Classify text blocks into appropriate styles."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1
        )
        
        styles = [{"id": s.id, "type": s.type, "content": s.content} for s in result.styles]
        print(f"LLM analyzed {len(styles)} blocks")
        
        return styles
    except Exception as e:
        print(f"Error analyzing styles with LLM: {e}")
        print(traceback.format_exc())
        return []


def save_styles(styles: List[Dict[str, Any]]) -> bool:
    """
    Save analyzed styles
    """
    try:
        print(f"Saving {len(styles)} styles...")
        response = requests.post(
            f"{SERVER_URL}/analyze-styles",
            json={"styles": styles}
        )
        
        if response.status_code == 200:
            print("Styles saved successfully")
            return True
        else:
            print(f"Failed to save styles: {response.json()}")
            return False
    except Exception as e:
        print(f"Error saving styles: {e}")
        print(traceback.format_exc())
        return False


def generate_document(styles: List[Dict[str, Any]]) -> bool:
    """
    Generate DOCX document from styles
    """
    try:
        print(f"Generating document from {len(styles)} styles...")
        response = requests.post(
            f"{SERVER_URL}/generate-document",
            json={"parsed_styles": styles}
        )
        
        if response.status_code == 200:
            data = response.json()
            print(f"Document generated successfully: {data.get('file_path')}")
            return True
        else:
            print(f"Failed to generate document: {response.json()}")
            return False
    except Exception as e:
        print(f"Error generating document: {e}")
        print(traceback.format_exc())
        return False


def main():
    """
    Main workflow for document processing
    """
    try:
        print("=" * 50)
        print("Starting document processing workflow")
        print("=" * 50)
        
        # Step 1: Parse DOCX file
        print("\n[Step 1] Parsing DOCX file...")
        if not parse_docx_file(FILE_URL):
            print("Failed to parse file, exiting...")
            return
        
        # Step 2: Get short blocks
        print("\n[Step 2] Getting short blocks...")
        short_blocks = get_short_blocks()
        if not short_blocks:
            print("No short blocks found, exiting...")
            return
        
        # Step 3: Analyze styles with LLM
        print("\n[Step 3] Analyzing styles with LLM...")
        styles = analyze_styles_with_llm(short_blocks)
        if not styles:
            print("Failed to analyze styles, exiting...")
            return
        
        # Step 4: Save styles
        print("\n[Step 4] Saving styles...")
        if not save_styles(styles):
            print("Failed to save styles, exiting...")
            return
        
        # Step 5: Generate document
        print("\n[Step 5] Generating document...")
        if not generate_document(styles):
            print("Failed to generate document, exiting...")
            return
        
        print("\n" + "=" * 50)
        print("Document processing completed successfully!")
        print("=" * 50)
        
    except Exception as e:
        print(f"Error in main workflow: {e}")
        print(traceback.format_exc())


if __name__ == '__main__':
    main()