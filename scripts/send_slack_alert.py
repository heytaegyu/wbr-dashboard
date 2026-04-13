#!/usr/bin/env python3
import argparse
import json
import sys
import urllib.error
import urllib.request
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
WEBHOOK_PATH = PROJECT_ROOT / ".codex-local" / "slack-webhook-url"
DEFAULT_URL = "https://wbr-dashboard.tracecorp.co.kr"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Send a WBR dashboard deployment alert to Slack.")
    parser.add_argument("--status", choices=["success", "failure"], required=True)
    parser.add_argument("--week", required=True, help="Week label such as W15")
    parser.add_argument("--summary", required=True, help="Short KPI summary")
    parser.add_argument("--details", default="", help="Optional extra detail")
    parser.add_argument("--url", default=DEFAULT_URL, help="Dashboard URL included in the Slack message")
    return parser.parse_args()


def read_webhook_url() -> str:
    if not WEBHOOK_PATH.exists():
        raise FileNotFoundError(f"Slack webhook file not found: {WEBHOOK_PATH}")
    return WEBHOOK_PATH.read_text(encoding="utf-8").strip()


def build_payload(args: argparse.Namespace) -> dict:
    status_emoji = ":white_check_mark:" if args.status == "success" else ":x:"
    status_label = "배포 완료" if args.status == "success" else "배포 실패"
    text = f"{status_emoji} WBR Dashboard 자동화 {status_label} - {args.week}"
    section_text = f"*{args.week}* {status_label}\n{args.summary}\n<{args.url}|wbr-dashboard.tracecorp.co.kr>"
    if args.details:
        section_text = f"{section_text}\n{args.details}"
    return {
        "text": text,
        "blocks": [
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": section_text,
                },
            }
        ],
    }


def send_payload(webhook_url: str, payload: dict) -> None:
    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        webhook_url,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=15) as response:
        result = response.read().decode("utf-8").strip()
        if response.status != 200 or result.lower() != "ok":
            raise RuntimeError(f"Slack webhook returned {response.status}: {result}")


def main() -> int:
    args = parse_args()
    try:
        webhook_url = read_webhook_url()
        payload = build_payload(args)
        send_payload(webhook_url, payload)
    except (OSError, urllib.error.URLError, RuntimeError) as exc:
        print(f"Failed to send Slack alert: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
