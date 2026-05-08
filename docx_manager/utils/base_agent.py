import os
import re
import json
import datetime
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

_LOG_DIR = Path(__file__).parent.parent.parent / 'logs'


def _log_llm_call(
    func_name: str,
    system_prompt: str,
    user_prompt: str,
    response_text: str,
    usage: dict,
) -> None:
    try:
        _LOG_DIR.mkdir(exist_ok=True)
        log_path = _LOG_DIR / f"llm_{datetime.date.today().isoformat()}.jsonl"
        entry = {
            "ts": datetime.datetime.now().isoformat(timespec="seconds"),
            "func": func_name,
            "system_prompt": system_prompt,
            "user_prompt": user_prompt,
            "response": response_text,
            "usage": usage,
        }
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception as e:
        print(f"[LLM LOG WARN] 日志写入失败: {e}")


def _extract_usage(usage_obj) -> dict:
    if usage_obj is None:
        return {}
    return {
        "prompt_tokens": getattr(usage_obj, "prompt_tokens", None),
        "completion_tokens": getattr(usage_obj, "completion_tokens", None),
        "total_tokens": getattr(usage_obj, "total_tokens", None),
    }


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
    text = _strip_thinking(response.choices[0].message.content)
    _log_llm_call("call", system_prompt, user_prompt, text, _extract_usage(response.usage))
    return text


def call_structured(system_prompt: str, user_prompt: str, response_model: Type[T]) -> T:
    model_result, completion = _get_instructor_client().chat.completions.create_with_completion(
        model=os.getenv('LLM_MODEL'),
        messages=[
            {'role': 'system', 'content': system_prompt},
            {'role': 'user', 'content': user_prompt},
        ],
        response_model=response_model,
        extra_body={'enable_thinking': False},
    )
    _log_llm_call(
        "call_structured",
        system_prompt,
        user_prompt,
        model_result.model_dump_json(ensure_ascii=False) if hasattr(model_result, "model_dump_json") else str(model_result),
        _extract_usage(getattr(completion, "usage", None)),
    )
    return model_result


def call_structured_with_image(
    system_prompt: str,
    user_prompt: str,
    base64_str: str,
    response_model: Type[T],
) -> T:
    model_result, completion = _get_vl_instructor_client().chat.completions.create_with_completion(
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
    _log_llm_call(
        "call_structured_with_image",
        system_prompt,
        f"[image base64 omitted] {user_prompt}",
        model_result.model_dump_json(ensure_ascii=False) if hasattr(model_result, "model_dump_json") else str(model_result),
        _extract_usage(getattr(completion, "usage", None)),
    )
    return model_result
