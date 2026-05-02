import requests


class BaseAgent:

    def __init__(self, endpoint: str, api_key: str = "", model: str = ""):
        self.endpoint = endpoint
        self.api_key = api_key
        self.model = model

    def _build_request(self, system_prompt: str, user_prompt: str) -> dict:
        # [待确认] 以平台实际请求格式为准
        payload = {
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user",   "content": user_prompt},
            ]
        }
        if self.model:
            payload["model"] = self.model
        return payload

    def _parse_response(self, response_json: dict) -> str:
        # [待确认] 以平台实际响应结构为准
        return response_json["choices"][0]["message"]["content"]

    def chat(self, user_prompt: str, system_prompt: str = "") -> str:
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        payload = self._build_request(system_prompt, user_prompt)

        try:
            resp = requests.post(self.endpoint, json=payload, headers=headers, timeout=60)
            resp.raise_for_status()
        except requests.exceptions.RequestException as e:
            raise ConnectionError(f"请求平台失败: {e}") from e

        try:
            return self._parse_response(resp.json())
        except (KeyError, IndexError, TypeError) as e:
            raise ValueError(f"响应结构不符合预期: {e}") from e
