"""调用大模型（默认 DeepSeek，任何 OpenAI 兼容接口都可以），要求返回 JSON。"""
import json
import time

import openai

import config

_client = None


def _get_client():
    global _client
    if _client is None:
        if not config.LLM_API_KEY:
            raise SystemExit("未找到 DEEPSEEK_API_KEY：请把 .env.example 复制为 .env，并填入你的 key。")
        _client = openai.OpenAI(api_key=config.LLM_API_KEY, base_url=config.LLM_BASE_URL)
    return _client


def ask_json(system: str, user: str, retries: int = 3) -> dict:
    """发送一次对话，返回模型输出的 JSON 对象。网络错误或 JSON 解析失败会自动重试。"""
    for attempt in range(retries):
        try:
            resp = _get_client().chat.completions.create(
                model=config.LLM_MODEL,
                messages=[{"role": "system", "content": system},
                          {"role": "user", "content": user}],
                response_format={"type": "json_object"},
                temperature=0,
            )
            return json.loads(resp.choices[0].message.content)
        except (json.JSONDecodeError, openai.APIConnectionError,
                openai.RateLimitError, openai.InternalServerError):
            if attempt == retries - 1:
                raise
            time.sleep(2 ** attempt)
