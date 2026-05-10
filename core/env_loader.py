"""환경 변수 / .env 파일 로더.

Single source of truth for environment variable loading.
Replaces 5+ duplicated parse_env_file() functions across modules.
"""

import os
from typing import Optional

DEFAULT_ENV_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "config",
    ".env",
)


def load_env(env_path: Optional[str] = None) -> dict:
    """`config/.env` 파일을 파싱합니다.

    지원 형식:
        KEY=value
        KEY="quoted value"
        KEY='single quoted'
        # 주석
        KEY=value  # 인라인 주석

    Args:
        env_path: .env 파일 경로 (기본: config/.env)

    Returns:
        {key: value} 딕셔너리. 파일이 없으면 빈 dict.
    """
    path = env_path or DEFAULT_ENV_PATH
    env_vars = {}

    if not os.path.exists(path):
        return env_vars

    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip()
                # 인라인 주석 제거 (따옴표 밖의 #만)
                if not value.startswith(("'", '"')) and "#" in value:
                    value = value.split("#")[0].strip()
                # 따옴표 제거
                if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
                    value = value[1:-1]
                env_vars[key] = value
    except OSError:
        pass

    return env_vars


def get_secret(key: str, env_path: Optional[str] = None,
               default: Optional[str] = None) -> Optional[str]:
    """환경변수 우선, .env 폴백으로 비밀값 조회.

    GitHub Actions 환경: os.environ 사용
    로컬 환경: config/.env 폴백

    Args:
        key: 변수명 (예: "TELEGRAM_FINANCE_BOT_TOKEN")
        env_path: .env 경로 (기본: config/.env)
        default: 못 찾았을 때 반환값

    Returns:
        변수값 또는 default
    """
    # 1순위: 환경변수 (GitHub Actions)
    val = os.environ.get(key)
    if val:
        return val

    # 2순위: .env 파일 (로컬)
    env_vars = load_env(env_path)
    return env_vars.get(key, default)
