"""Shared helpers for interacting with the Kimi (Moonshot) API."""
import os
from typing import Any, Dict, List

from openai import OpenAI

DEFAULT_MODEL = "kimi-k2-turbo-preview"
DEFAULT_BASE_URL = "https://api.moonshot.cn/v1"


def build_client(api_key: str | None = None, base_url: str | None = None) -> OpenAI:
    return OpenAI(
        api_key=api_key or os.getenv("MOONSHOT_API_KEY") or "sk-PK4uOTEqNwfrVmKsP6ew5lMFmPtfqIN9uJMEKHED52VnsiPy",
        base_url=base_url or os.getenv("MOONSHOT_API_BASE") or DEFAULT_BASE_URL,
    )


def list_models(client: OpenAI) -> List[str]:
    models = client.models.list()
    return [m.id for m in models.data]


def chat_completion(client: OpenAI, messages: List[Dict[str, Any]], model: str | None = None, **kwargs):
    return client.chat.completions.create(
        model=model or DEFAULT_MODEL,
        messages=messages,
        **kwargs,
    )


__all__ = ["build_client", "list_models", "chat_completion", "DEFAULT_MODEL", "DEFAULT_BASE_URL"]
