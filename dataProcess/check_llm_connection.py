#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Check TextProcessor connectivity only.

Legacy LLM_API_URL/api/tags probing is intentionally removed.
"""

import os
import requests
from dotenv import load_dotenv


def _build_chat_payload(provider: str, model: str) -> dict:
    return {
        "prompt": "ping",
        "provider": provider,
        "model": model,
        "max_tokens": 32,
        "temperature": 0.1,
    }


def check_textprocessor_chat() -> int:
    load_dotenv()

    textprocessor_url = os.getenv("TEXTPROCESSOR_URL", "http://10.1.3.127:6017/chat").strip()
    textprocessor_provider = os.getenv("TEXTPROCESSOR_PROVIDER", "remote").strip()
    textprocessor_model = os.getenv("TEXTPROCESSOR_MODEL", "remote8b").strip()

    print("=== TextProcessor /chat connection check ===")
    print(f"TEXTPROCESSOR_URL: {textprocessor_url}")
    print(f"TEXTPROCESSOR_PROVIDER: {textprocessor_provider}")
    print(f"TEXTPROCESSOR_MODEL: {textprocessor_model}")

    # Health endpoint is derived from /chat URL for quick service status verification.
    health_url = textprocessor_url.replace("/chat", "/health")
    try:
        health_resp = requests.get(health_url, timeout=5)
        print(f"[health] status_code={health_resp.status_code}")
        if health_resp.status_code == 200:
            print(f"[health] body={health_resp.text}")
        else:
            print(f"[health] body={health_resp.text}")
    except Exception as e:
        print(f"[health] request failed: {e}")

    payload = _build_chat_payload(textprocessor_provider, textprocessor_model)
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "TextProcessor-Client/1.0",
    }

    print(f"[chat] POST {textprocessor_url}")
    print(f"[chat] payload={payload}")

    try:
        resp = requests.post(textprocessor_url, json=payload, headers=headers, timeout=30)
    except requests.exceptions.Timeout:
        print("[chat] request timeout")
        return 1
    except requests.exceptions.ConnectionError as e:
        print(f"[chat] connection error: {e}")
        return 1
    except Exception as e:
        print(f"[chat] unexpected error: {e}")
        return 1

    print(f"[chat] status_code={resp.status_code}")
    if resp.status_code != 200:
        print(f"[chat] body={resp.text}")
        return 1

    try:
        data = resp.json()
    except Exception as e:
        print(f"[chat] invalid json response: {e}")
        print(f"[chat] body={resp.text}")
        return 1

    output = data.get("output")
    provider = data.get("provider")
    model_alias = data.get("model_alias")
    post_id = data.get("post_id")

    print("[chat] success")
    print(f"[chat] provider={provider}")
    print(f"[chat] model_alias={model_alias}")
    print(f"[chat] post_id={post_id}")
    print(f"[chat] output={output}")
    return 0


def main() -> None:
    exit_code = check_textprocessor_chat()
    raise SystemExit(exit_code)


if __name__ == "__main__":
    main()
