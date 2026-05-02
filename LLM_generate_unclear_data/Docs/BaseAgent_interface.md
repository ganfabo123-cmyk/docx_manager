# BaseAgent 接口文档

## 1. 模块概述

**文件路径：** `LLM_generate_unclear_data/base_agent.py`  
**职责：** 封装与远程大模型平台的 HTTP 通信，对外暴露统一的 `chat()` 接口。

---

## 2. 类：`BaseAgent`

### 2.1 构造函数

```python
BaseAgent(endpoint: str, api_key: str = "", model: str = "")
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `endpoint` | `str` | 是 | 远程平台的 HTTP 接口地址 |
| `api_key` | `str` | 否 | 鉴权密钥，默认为空字符串 |
| `model` | `str` | 否 | 指定模型名称，默认为空字符串（使用平台默认） |

---

### 2.2 公开方法

#### `chat(user_prompt, system_prompt="") → str`

对外唯一调用接口。

```python
def chat(self, user_prompt: str, system_prompt: str = "") -> str
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `user_prompt` | `str` | 用户侧提示词 |
| `system_prompt` | `str` | 系统提示词，可选，默认为空 |

**返回值：** `str` — 模型回复的纯文本内容

**异常：**
- `ConnectionError` — HTTP 请求失败或超时
- `ValueError` — 响应结构不符合预期，无法提取回复文本

---

### 2.3 私有方法

#### `_build_request(system_prompt, user_prompt) → dict`

将 prompt 组装为平台要求的请求体。

```python
def _build_request(self, system_prompt: str, user_prompt: str) -> dict
```

**返回值：** `dict` — 请求体，格式如下（**[待确认]** 以实际平台为准）：

```json
{
  "model": "<model_name>",
  "messages": [
    {"role": "system", "content": "<system_prompt>"},
    {"role": "user",   "content": "<user_prompt>"}
  ]
}
```

> **[待确认]** 平台的请求体字段名、消息结构、是否需要额外字段（如 `temperature`、`max_tokens`）。

---

#### `_parse_response(response_json) → str`

从平台响应中提取模型回复文本。

```python
def _parse_response(self, response_json: dict) -> str
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `response_json` | `dict` | 平台返回的响应体（已解析为 dict） |

**返回值：** `str` — 模型回复的纯文本

**预期响应结构（[待确认]）：**

```json
{
  "choices": [
    {
      "message": {
        "content": "<模型回复文本>"
      }
    }
  ]
}
```

> **[待确认]** 平台实际响应字段名及层级结构。

---

## 3. 调用流程

```
chat(user_prompt, system_prompt)
    │
    ├─ _build_request()  →  构造请求体 dict
    │
    ├─ HTTP POST to endpoint  →  携带 api_key 鉴权头
    │       │
    │       └─ 失败 → 抛出 ConnectionError
    │
    ├─ _parse_response()  →  提取回复文本
    │       │
    │       └─ 结构异常 → 抛出 ValueError
    │
    └─ 返回 str
```

---

## 4. 鉴权方式

> **[待确认]** 以下为常见约定，以实际平台要求为准。

预设方案：在请求头中携带 API Key：

```
Authorization: Bearer <api_key>
```

---

## 5. 使用示例

```python
agent = BaseAgent(
    endpoint="https://platform.example.com/v1/chat",
    api_key="your-api-key",
    model="your-model-name"
)

reply = agent.chat(
    user_prompt="将以下公式转换为标准 LaTeX：E=mc^2",
    system_prompt="你是一个数学公式转换专家。"
)
print(reply)
```
