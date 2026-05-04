import os
import re
from pathlib import Path
from typing import TypeVar, Type
from dotenv import load_dotenv
from openai import OpenAI
from pydantic import BaseModel
import instructor

load_dotenv(Path(__file__).parent.parent.parent / '.env')

T = TypeVar('T', bound=BaseModel)

_raw_client: OpenAI | None = None
_instructor_client = None
_vl_instructor_client = None


def _get_raw_client() -> OpenAI:
    global _raw_client
    if _raw_client is None:
        _raw_client = OpenAI(
            api_key=os.getenv('LLM_API_KEY'),
            base_url=os.getenv('LLM_API_BASE_URL'),
        )
    return _raw_client


def _get_instructor_client():
    global _instructor_client
    if _instructor_client is None:
        _instructor_client = instructor.from_openai(_get_raw_client())
    return _instructor_client


def _get_vl_instructor_client():
    global _vl_instructor_client
    if _vl_instructor_client is None:
        vl_raw = OpenAI(
            api_key=os.getenv('LLM_API_KEY'),
            base_url=os.getenv('LLM_API_BASE_URL'),
        )
        _vl_instructor_client = instructor.from_openai(vl_raw)
    return _vl_instructor_client


def _strip_thinking(text: str) -> str:
    return re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL).strip()


def call(system_prompt: str, user_prompt: str) -> str:
    response = _get_raw_client().chat.completions.create(
        model=os.getenv('LLM_MODEL'),
        messages=[
            {'role': 'system', 'content': system_prompt},
            {'role': 'user', 'content': user_prompt},
        ],
        extra_body={'enable_thinking': False},
    )
    return _strip_thinking(response.choices[0].message.content)


def call_structured(system_prompt: str, user_prompt: str, response_model: Type[T]) -> T:
    return _get_instructor_client().chat.completions.create(
        model=os.getenv('LLM_MODEL'),
        messages=[
            {'role': 'system', 'content': system_prompt},
            {'role': 'user', 'content': user_prompt},
        ],
        response_model=response_model,
        extra_body={'enable_thinking': False},
    )


def call_structured_with_image(
    system_prompt: str,
    user_prompt: str,
    base64_str: str,
    response_model: Type[T],
) -> T:
    return _get_vl_instructor_client().chat.completions.create(
        model=os.getenv('LLM_VL_NODEL'),
        messages=[
            {'role': 'system', 'content': system_prompt},
            {
                'role': 'user',
                'content': [
                    {
                        'type': 'image_url',
                        'image_url': {'url': f'data:image/png;base64,{base64_str}'},
                    },
                    {'type': 'text', 'text': user_prompt},
                ],
            },
        ],
        response_model=response_model,
        extra_body={'enable_thinking': False},
    )
