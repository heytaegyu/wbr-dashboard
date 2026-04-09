#!/usr/bin/env python3
import json
import os
import sys
import urllib.error
import urllib.request


MESSAGES = {
    "morning": "오전: 오늘 꼭 끝낼 1개는?",
    "afternoon": "오후: 지금 제일 중요한 일 하는 중인가?",
    "evening": "저녁: 오늘 진짜 해낸 게 있나?",
}


def require_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def build_payload(slot: str) -> dict:
    text = MESSAGES[slot]
    return {
        "channel": require_env("SLACK_CHANNEL_ID"),
        "text": text,
    }


def send_message(slot: str) -> None:
    token = require_env("SLACK_BOT_TOKEN")
    payload = json.dumps(build_payload(slot)).encode("utf-8")
    request = urllib.request.Request(
        "https://slack.com/api/chat.postMessage",
        data=payload,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=utf-8",
        },
        method="POST",
    )

    with urllib.request.urlopen(request, timeout=20) as response:
        body = response.read().decode("utf-8")
        data = json.loads(body)
        if response.status != 200 or not data.get("ok"):
            raise RuntimeError(f"Slack API error: {data}")


def main() -> int:
    if len(sys.argv) != 2 or sys.argv[1] not in MESSAGES:
        print("Usage: send_ceo_prompt.py [morning|afternoon|evening]", file=sys.stderr)
        return 1

    try:
        send_message(sys.argv[1])
    except (KeyError, OSError, urllib.error.URLError, RuntimeError, json.JSONDecodeError) as exc:
        print(f"Failed to send CEO prompt: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
