"""yfinance 헬퍼: 티커 해석 + 재시도 로직.

Single source of truth for yfinance interactions.
Replaces 3+ duplicated .KS/.KQ resolution logic.
"""

import time
from typing import Optional

try:
    import yfinance as yf
    import pandas as pd
except ImportError:
    yf = None
    pd = None


def resolve_ticker(raw: str) -> str:
    """원시 ticker → yfinance 호환 ticker로 해석.

    규칙:
    - 6자리 숫자 (한국 주식): .KS 우선 시도, 데이터 없으면 .KQ
    - 그 외: 변경 없음

    Args:
        raw: 원시 티커 (예: "005930", "AAPL", "BTC-USD")

    Returns:
        yfinance 호환 티커 (예: "005930.KS")
    """
    if not raw:
        return raw

    # 이미 접미사가 있는 경우
    if "." in raw or "-" in raw:
        return raw

    # 6자리 숫자 = 한국 주식
    if raw.isdigit() and len(raw) == 6:
        if yf is None:
            return f"{raw}.KS"  # 라이브러리 없으면 KS 추정

        # KS 우선 → 실패 시 KQ
        for suffix in [".KS", ".KQ"]:
            try:
                t = yf.Ticker(f"{raw}{suffix}")
                hist = t.history(period="5d")
                if hist is not None and not hist.empty and len(hist) >= 1:
                    info = t.info or {}
                    name = info.get("longName") or info.get("shortName")
                    if name and name != raw:
                        return f"{raw}{suffix}"
            except Exception:
                continue
        return f"{raw}.KS"  # 폴백

    return raw


def fetch_history(ticker: str, period: str = "1y", retries: int = 3,
                  backoff: float = 2.0) -> Optional["pd.DataFrame"]:
    """yfinance history with 재시도 + 지수 백오프.

    Args:
        ticker: 티커 (이미 해석됨)
        period: 기간 (1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max)
        retries: 최대 재시도 횟수
        backoff: 백오프 배수 (초)

    Returns:
        DataFrame 또는 None (모두 실패)
    """
    if yf is None:
        return None

    last_err = None
    for attempt in range(retries):
        try:
            t = yf.Ticker(ticker)
            hist = t.history(period=period)
            if hist is not None and not hist.empty:
                return hist
        except Exception as e:
            last_err = e

        if attempt < retries - 1:
            time.sleep(backoff ** attempt)  # 1, 2, 4초

    return None


def fetch_info_safely(ticker: str, retries: int = 2, backoff: float = 2.0) -> dict:
    """yfinance info 안전 조회.

    Args:
        ticker: 티커
        retries: 재시도 횟수
        backoff: 백오프

    Returns:
        info dict (실패 시 빈 dict)
    """
    if yf is None:
        return {}

    for attempt in range(retries):
        try:
            t = yf.Ticker(ticker)
            info = t.info
            if info:
                return info
        except Exception:
            pass

        if attempt < retries - 1:
            time.sleep(backoff ** attempt)

    return {}


def get_current_price(ticker: str) -> Optional[float]:
    """info.currentPrice 우선, regularMarketPrice 폴백.

    history()의 NaN 문제를 회피하기 위해 info 사용.

    Args:
        ticker: 티커

    Returns:
        현재가 또는 None
    """
    info = fetch_info_safely(ticker)
    price = info.get("currentPrice") or info.get("regularMarketPrice")
    if price is None:
        return None
    try:
        return float(price)
    except (ValueError, TypeError):
        return None
