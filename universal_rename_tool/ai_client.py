import json
import re
from dataclasses import dataclass

import requests

from .config import DEFAULT_AI_BASE_URL, DEFAULT_AI_MODEL


@dataclass
class AIConfig:
    enabled: bool = False
    api_key: str = ""
    base_url: str = DEFAULT_AI_BASE_URL
    model: str = DEFAULT_AI_MODEL
    timeout: int = 30
    temperature: float = 0.2


class DSV4FlashClient:
    def __init__(self, config: AIConfig):
        self.config = config

    def _endpoint(self):
        base = (self.config.base_url or DEFAULT_AI_BASE_URL).rstrip("/")
        if base.endswith("/chat/completions"):
            return base
        return f"{base}/chat/completions"

    def _headers(self):
        return {"Authorization": f"Bearer {self.config.api_key}", "Content-Type": "application/json"}

    def test_connection(self) -> dict:
        if not self.config.enabled:
            return {"ok": False, "data": None, "raw_text": "", "error": "请先启用 AI。"}
        if not self.config.api_key:
            return {"ok": False, "data": None, "raw_text": "", "error": "请先填写 API Key。"}
        return self.chat_json([
            {"role": "system", "content": "Return strict JSON only."},
            {"role": "user", "content": '{"ping":"pong"}'},
        ], temperature=0, timeout=min(self.config.timeout, 15))

    def chat_json(self, messages, temperature=None, timeout=None) -> dict:
        if not self.config.api_key:
            return {"ok": False, "data": None, "raw_text": "", "error": "请先填写 API Key。"}
        payload = {
            "model": self.config.model or DEFAULT_AI_MODEL,
            "messages": messages,
            "temperature": self.config.temperature if temperature is None else temperature,
        }
        try:
            response = requests.post(
                self._endpoint(),
                headers=self._headers(),
                json=payload,
                timeout=timeout or self.config.timeout,
            )
            if response.status_code in {401, 403}:
                return {"ok": False, "data": None, "raw_text": "", "error": "认证失败，请检查 API Key 或接口权限。"}
            if response.status_code == 429:
                return {"ok": False, "data": None, "raw_text": "", "error": "请求过于频繁或额度不足。"}
            if response.status_code >= 500:
                return {"ok": False, "data": None, "raw_text": "", "error": "AI 服务暂时不可用，请稍后重试。"}
            response.raise_for_status()
            body = response.json()
            raw_text = body.get("choices", [{}])[0].get("message", {}).get("content", "")
            data = self._parse_json_text(raw_text)
            return {"ok": True, "data": data, "raw_text": raw_text, "error": ""}
        except requests.Timeout:
            return {"ok": False, "data": None, "raw_text": "", "error": "AI 请求超时。"}
        except requests.RequestException as exc:
            return {"ok": False, "data": None, "raw_text": "", "error": f"AI 请求失败：{type(exc).__name__}"}
        except Exception as exc:
            return {"ok": False, "data": None, "raw_text": "", "error": f"AI 响应解析失败：{type(exc).__name__}"}

    @staticmethod
    def _parse_json_text(text):
        text = str(text or "").strip()
        if not text:
            raise ValueError("empty response")
        fenced = re.search(r"```(?:json)?\s*(.*?)```", text, re.S | re.I)
        if fenced:
            text = fenced.group(1).strip()
        return json.loads(text)
