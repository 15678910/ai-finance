"""텔레그램 봇 메시징.

Single source of truth for Telegram API.
Replaces 11+ duplicated send_telegram() functions across modules.
"""

import os
import json
import urllib.request
import urllib.parse
import urllib.error
from typing import Optional

from .env_loader import get_secret

TELEGRAM_API_BASE = "https://api.telegram.org"
MESSAGE_MAX_LENGTH = 4096  # Telegram limit
DOCUMENT_MAX_SIZE = 50 * 1024 * 1024  # 50 MB


def _get_credentials(bot_token: Optional[str] = None,
                     chat_id: Optional[str] = None) -> tuple[Optional[str], Optional[str]]:
    """봇 토큰과 chat_id를 환경/.env에서 조회."""
    bot_token = bot_token or get_secret("TELEGRAM_FINANCE_BOT_TOKEN")
    chat_id = chat_id or get_secret("TELEGRAM_FINANCE_CHAT_ID")
    return bot_token, chat_id


def send_message(text: str, bot_token: Optional[str] = None,
                 chat_id: Optional[str] = None, timeout: int = 10,
                 disable_preview: bool = True) -> bool:
    """텔레그램 텍스트 메시지 전송.

    Args:
        text: 메시지 본문 (4096자 초과 시 자동 절단)
        bot_token: 봇 토큰 (생략 시 환경변수)
        chat_id: 채팅 ID (생략 시 환경변수)
        timeout: HTTP 타임아웃 (초)
        disable_preview: URL 프리뷰 비활성화 여부

    Returns:
        성공 여부
    """
    bot_token, chat_id = _get_credentials(bot_token, chat_id)
    if not bot_token or not chat_id:
        return False

    # 메시지 길이 제한
    if len(text) > MESSAGE_MAX_LENGTH:
        text = text[:MESSAGE_MAX_LENGTH - 50] + "\n\n... (잘림)"

    url = f"{TELEGRAM_API_BASE}/bot{bot_token}/sendMessage"
    data = urllib.parse.urlencode({
        "chat_id": chat_id,
        "text": text,
        "disable_web_page_preview": "true" if disable_preview else "false",
    }).encode("utf-8")

    try:
        req = urllib.request.Request(url, data=data, method="POST")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return bool(result.get("ok"))
    except (urllib.error.URLError, urllib.error.HTTPError, json.JSONDecodeError, OSError):
        return False


def send_document(file_path: str, filename: Optional[str] = None,
                  caption: Optional[str] = None,
                  bot_token: Optional[str] = None,
                  chat_id: Optional[str] = None,
                  timeout: int = 30) -> bool:
    """텔레그램 파일 전송 (multipart/form-data).

    Args:
        file_path: 전송할 파일 경로
        filename: 표시할 파일명 (생략 시 basename)
        caption: 파일 캡션
        bot_token: 봇 토큰
        chat_id: 채팅 ID
        timeout: 타임아웃

    Returns:
        성공 여부
    """
    bot_token, chat_id = _get_credentials(bot_token, chat_id)
    if not bot_token or not chat_id or not os.path.exists(file_path):
        return False

    # 파일 크기 체크
    if os.path.getsize(file_path) > DOCUMENT_MAX_SIZE:
        return False

    if not filename:
        filename = os.path.basename(file_path)

    url = f"{TELEGRAM_API_BASE}/bot{bot_token}/sendDocument"
    boundary = "----PythonFormBoundary"

    try:
        with open(file_path, "rb") as f:
            file_data = f.read()
    except OSError:
        return False

    # multipart/form-data 본문 조립
    parts = []
    parts.append(f"--{boundary}".encode())
    parts.append(b'Content-Disposition: form-data; name="chat_id"\r\n')
    parts.append(str(chat_id).encode() + b"\r\n")

    if caption:
        parts.append(f"--{boundary}".encode())
        parts.append(b'Content-Disposition: form-data; name="caption"\r\n')
        parts.append(caption.encode("utf-8") + b"\r\n")

    parts.append(f"--{boundary}".encode())
    parts.append(
        f'Content-Disposition: form-data; name="document"; filename="{filename}"\r\n'.encode("utf-8")
    )
    parts.append(b"Content-Type: application/octet-stream\r\n\r\n")
    parts.append(file_data)
    parts.append(f"\r\n--{boundary}--\r\n".encode())

    body = b"\r\n".join(parts)

    try:
        req = urllib.request.Request(url, data=body, method="POST")
        req.add_header("Content-Type", f"multipart/form-data; boundary={boundary}")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return bool(result.get("ok"))
    except (urllib.error.URLError, urllib.error.HTTPError, json.JSONDecodeError, OSError):
        return False
