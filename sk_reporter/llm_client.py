"""Ollama: локальный daemon или облако ollama.com (OLLAMA_API_KEY)."""

from __future__ import annotations

import os
from functools import lru_cache
from typing import Any

DEFAULT_MODEL = "gemma4:31b-cloud"
OLLAMA_CLOUD_HOST = "https://ollama.com"
OLLAMA_LOCAL_HOST = "http://127.0.0.1:11434"


def default_model() -> str:
    return (os.getenv("OLLAMA_MODEL") or DEFAULT_MODEL).strip() or DEFAULT_MODEL


def llm_status() -> dict[str, Any]:
    api_key = (os.getenv("OLLAMA_API_KEY") or "").strip()
    host_override = (os.getenv("OLLAMA_HOST") or "").strip()
    if api_key:
        mode = "cloud"
        host = host_override or OLLAMA_CLOUD_HOST
    else:
        mode = "local"
        host = host_override or OLLAMA_LOCAL_HOST
    return {
        "mode": mode,
        "host": host,
        "model": default_model(),
        "api_key_set": bool(api_key),
    }


@lru_cache(maxsize=1)
def _get_client():
    from ollama import Client

    status = llm_status()
    headers: dict[str, str] | None = None
    api_key = (os.getenv("OLLAMA_API_KEY") or "").strip()
    if api_key:
        headers = {"Authorization": f"Bearer {api_key}"}
    return Client(host=status["host"], headers=headers)


def _as_dict(response: Any) -> dict[str, Any]:
    if isinstance(response, dict):
        return response
    if hasattr(response, "model_dump"):
        return response.model_dump()
    message = getattr(response, "message", None)
    content = getattr(message, "content", "") if message else ""
    return {"message": {"content": content or ""}}


def llm_chat(
    *,
    model: str | None = None,
    messages: list[dict[str, str]],
    stream: bool = False,
    options: dict[str, Any] | None = None,
) -> dict[str, Any]:
    kwargs: dict[str, Any] = {
        "model": model or default_model(),
        "messages": messages,
        "stream": stream,
    }
    if options:
        kwargs["options"] = options
    response = _get_client().chat(**kwargs)
    return _as_dict(response)


def ping_llm() -> tuple[bool, str]:
    """Проверка доступности Ollama (list models)."""
    try:
        client = _get_client()
        client.list()
        return True, "ok"
    except Exception as exc:
        return False, str(exc)
