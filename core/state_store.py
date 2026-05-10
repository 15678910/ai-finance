"""JSON 상태 파일 저장소.

Single source of truth for state file management.
Replaces 6+ duplicated load_state()/save_state() pairs.

용도: 알림 중복 방지, 진행 상태 추적, 분석 히스토리 등.
"""

import os
import json
from datetime import datetime, timezone, timedelta
from typing import Optional, Any

DEFAULT_STATE_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "docs",
)


def _state_path(name: str, state_dir: Optional[str] = None) -> str:
    """상태 파일 경로 계산. name → docs/{name}_state.json"""
    base = state_dir or DEFAULT_STATE_DIR
    if not name.endswith(".json"):
        name = f"{name}_state.json"
    return os.path.join(base, name)


def load_state(name: str, default: Optional[dict] = None,
               state_dir: Optional[str] = None) -> dict:
    """상태 파일 로드.

    Args:
        name: 상태 이름 (예: "breaking_news") → docs/breaking_news_state.json
        default: 파일 없을 때 기본값 (None이면 {})
        state_dir: 디렉터리 (기본: docs/)

    Returns:
        상태 dict (오류 시 default)
    """
    path = _state_path(name, state_dir)
    fallback = default if default is not None else {}

    if not os.path.exists(path):
        return fallback

    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        # 손상된 파일은 default 반환 (silent corruption 방지를 위해 로그 권장)
        return fallback


def save_state(name: str, data: dict, state_dir: Optional[str] = None) -> bool:
    """상태 파일 저장.

    Args:
        name: 상태 이름
        data: 저장할 dict
        state_dir: 디렉터리

    Returns:
        성공 여부
    """
    path = _state_path(name, state_dir)
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)
        return True
    except (OSError, TypeError):
        return False


def is_recent_alert(state: dict, key: str, hours: int = 12) -> bool:
    """특정 알림이 최근 N시간 내 발송되었는지 확인.

    Args:
        state: load_state()로 받은 dict
        key: 알림 식별자
        hours: 임계 시간

    Returns:
        최근에 발송됨 → True (다시 보내지 말 것)
    """
    last_alerts = state.get("last_alerts", {}) if isinstance(state, dict) else {}
    last = last_alerts.get(key)
    if not last:
        return False

    try:
        last_dt = datetime.fromisoformat(last)
        if last_dt.tzinfo is None:
            last_dt = last_dt.replace(tzinfo=timezone.utc)
        now = datetime.now(timezone.utc)
        return (now - last_dt) < timedelta(hours=hours)
    except (ValueError, TypeError):
        return False


def mark_alert_sent(state: dict, key: str) -> dict:
    """알림 발송 시각 기록.

    Args:
        state: load_state()로 받은 dict (in-place 수정)
        key: 알림 식별자

    Returns:
        수정된 state
    """
    if not isinstance(state, dict):
        state = {}
    state.setdefault("last_alerts", {})[key] = datetime.now(timezone.utc).isoformat()
    return state
