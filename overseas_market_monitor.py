"""
해외 시장 실시간 모니터링 시스템
=================================

KOSPI 마감 후(15:30) 다음날 개장(09:00) 전까지의 해외 시장 흐름을 추적하고,
큰 변동(±2% 이상) 발생 시 텔레그램으로 즉시 알립니다.

수집 대상:
  - 미국: S&P 500, Nasdaq, Dow, VIX, NQ/ES 선물
  - 아시아: Nikkei (일본), Hang Seng (홍콩), Shanghai (중국)
  - 유럽: FTSE (영국), DAX (독일)
  - 통화/금리: 달러 인덱스, 미국 10년 국채

알림 조건:
  - 주요 지수 ±2% 이상 변동
  - VIX 30 초과 (시장 공포)
  - 미국 선물 ±1% 이상 (KOSPI 시초가 영향)

🚨 시뮬레이션/분석용. 자동 매매 절대 금지.
"""

import os
import sys
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta
from pathlib import Path

from core import send_message, load_state, save_state, is_recent_alert, mark_alert_sent

try:
    import yfinance as yf
except ImportError:
    print("[오류] yfinance 미설치. pip install yfinance")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "overseas_market.json")

KST = timezone(timedelta(hours=9))

# ====================================================================
# 모니터링 대상
# ====================================================================
MARKETS = {
    "미국_지수": [
        {"ticker": "^GSPC", "name": "S&P 500", "region": "🇺🇸"},
        {"ticker": "^IXIC", "name": "Nasdaq", "region": "🇺🇸"},
        {"ticker": "^DJI", "name": "Dow Jones", "region": "🇺🇸"},
        {"ticker": "^VIX", "name": "VIX", "region": "🇺🇸", "is_vix": True},
    ],
    "미국_선물": [
        {"ticker": "ES=F", "name": "S&P 선물", "region": "🇺🇸", "is_futures": True},
        {"ticker": "NQ=F", "name": "Nasdaq 선물", "region": "🇺🇸", "is_futures": True},
        {"ticker": "YM=F", "name": "Dow 선물", "region": "🇺🇸", "is_futures": True},
    ],
    "아시아": [
        {"ticker": "^N225", "name": "Nikkei 225", "region": "🇯🇵"},
        {"ticker": "^HSI", "name": "Hang Seng", "region": "🇭🇰"},
        {"ticker": "000001.SS", "name": "Shanghai", "region": "🇨🇳"},
    ],
    "유럽": [
        {"ticker": "^FTSE", "name": "FTSE 100", "region": "🇬🇧"},
        {"ticker": "^GDAXI", "name": "DAX", "region": "🇩🇪"},
    ],
    "통화_금리": [
        {"ticker": "DX-Y.NYB", "name": "달러 인덱스", "region": "🌍"},
        {"ticker": "^TNX", "name": "美 10Y 국채", "region": "🇺🇸"},
        {"ticker": "KRW=X", "name": "USD/KRW", "region": "🇰🇷"},
    ],
}


# ====================================================================
# 시장 운영 시간 (KST)
# ====================================================================
def get_market_hours_kst():
    """각 시장의 KST 운영 시간."""
    return {
        "KOSPI": (9, 15.5),       # 09:00~15:30
        "Nikkei": (9, 15),        # 09:00~15:00
        "HangSeng": (10.5, 16),   # 10:30~16:00
        "Shanghai": (10.5, 16),   # 10:30~16:00
        "FTSE": (17, 25.5),       # 17:00~01:30 (다음날)
        "DAX": (17, 25.5),
        "US_NYSE": (22.5, 29),    # 22:30~05:00 (서머타임 기준)
        # 일반 시간(11~3월)은 23:30~06:00이지만 단순화
    }


def is_kospi_closed():
    """KOSPI가 닫혀있는지 확인."""
    now = datetime.now(KST)
    if now.weekday() >= 5:  # 토/일
        return True
    hour = now.hour + now.minute / 60
    return hour < 9 or hour >= 15.5


def get_active_market():
    """현재 활성 주요 시장 (KST 기준)."""
    now = datetime.now(KST)
    hour = now.hour + now.minute / 60

    if now.weekday() >= 5:
        return "주말 휴장"
    if 9 <= hour < 15.5:
        return "🇰🇷 KOSPI 운영 중"
    if 22.5 <= hour < 29 or hour < 5:  # 미국 정규장
        return "🇺🇸 미국 시장 운영 중"
    if 17 <= hour < 25.5 or hour < 1.5:
        return "🇪🇺 유럽 시장 운영 중"
    return "전 세계 휴장 (또는 시간 외)"


# ====================================================================
# 데이터 수집
# ====================================================================
def fetch_market_data(ticker_info: dict) -> dict:
    """단일 시장 데이터 수집.

    fast_info (실시간 quote) → history() fallback 순서로 시도.
    history(period="5d")는 선물 티커에서 당일 미완성 바와 전일 바가
    동일 종가를 반환하는 경우가 있어 0.00% 오류 발생 → fast_info 우선 사용.
    """
    ticker = ticker_info["ticker"]
    try:
        stock = yf.Ticker(ticker)

        current: float | None = None
        prev: float | None = None
        change_pct: float = 0.0

        # ── 1) fast_info: 실시간 last_price + previous_close ──────────────
        try:
            fi = stock.fast_info
            _cur = fi.last_price
            _prv = fi.previous_close
            if _cur and _prv and float(_prv) > 0 and abs(float(_cur) - float(_prv)) > 1e-6:
                current = float(_cur)
                prev = float(_prv)
                change_pct = (current - prev) / prev * 100
        except Exception:
            pass  # fast_info 미지원 시 fallback

        # ── 2) history() fallback ─────────────────────────────────────────
        if current is None:
            hist = stock.history(period="5d")
            if hist.empty or len(hist) < 2:
                return None
            current = float(hist["Close"].iloc[-1])
            prev    = float(hist["Close"].iloc[-2])
            change_pct = (current - prev) / prev * 100 if prev and prev > 0 else 0.0

        # ── 3) 52주 고저 ───────────────────────────────────────────────────
        high_52w: float | None = None
        low_52w:  float | None = None
        try:
            hist_year = stock.history(period="1y")
            if not hist_year.empty:
                high_52w = round(float(hist_year["High"].max()), 2)
                low_52w  = round(float(hist_year["Low"].min()),  2)
        except Exception:
            pass

        return {
            "ticker":    ticker,
            "name":      ticker_info["name"],
            "region":    ticker_info.get("region", ""),
            "current":   round(current, 2),
            "previous":  round(prev, 2) if prev else None,
            "change_pct": round(change_pct, 2),
            "high_52w":  high_52w,
            "low_52w":   low_52w,
            "is_vix":    ticker_info.get("is_vix", False),
            "is_futures": ticker_info.get("is_futures", False),
        }
    except Exception as e:
        print(f"  [실패] {ticker_info['name']}: {e}")
        return None


# ====================================================================
# 알림 조건 판단
# ====================================================================
def should_alert(market: dict) -> tuple:
    """알림 발송 조건. (필요여부, 심각도, 메시지)"""
    name = market["name"]
    change = market["change_pct"]
    abs_change = abs(change)
    current = market["current"]

    # VIX 특수 처리
    if market.get("is_vix"):
        if current > 30:
            return True, "긴급", f"🔴 VIX {current} - 시장 공포 (30 초과)"
        elif current > 25 and abs_change > 10:
            return True, "경고", f"🟠 VIX {current} ({change:+.1f}%) - 변동성 급증"
        return False, "", ""

    # 선물은 ±1% 이상으로 KOSPI 시초가 영향
    if market.get("is_futures"):
        if abs_change >= 2:
            sign = "급등" if change > 0 else "급락"
            return True, "경고", f"🟠 {name} {sign} {change:+.2f}% - KOSPI 시초가 영향 예상"
        return False, "", ""

    # 일반 지수: ±2% 이상
    if abs_change >= 3:
        sign = "급등" if change > 0 else "급락"
        return True, "긴급", f"🔴 {name} {sign} {change:+.2f}%"
    elif abs_change >= 2:
        sign = "상승" if change > 0 else "하락"
        return True, "경고", f"🟠 {name} {sign} {change:+.2f}%"

    return False, "", ""


def predict_kospi_open(markets: list) -> str:
    """미국 선물 기반 KOSPI 시초가 예측."""
    futures = [m for m in markets if m.get("is_futures")]
    if not futures:
        return ""

    avg_change = sum(m["change_pct"] for m in futures) / len(futures)
    abs_avg = abs(avg_change)

    if abs_avg < 0.3:
        return "보합 시초가 예상"
    elif avg_change > 1.5:
        return f"강한 갭상승 시초가 예상 ({avg_change:+.2f}%)"
    elif avg_change > 0.5:
        return f"갭상승 시초가 예상 ({avg_change:+.2f}%)"
    elif avg_change < -1.5:
        return f"강한 갭하락 시초가 예상 ({avg_change:+.2f}%)"
    elif avg_change < -0.5:
        return f"갭하락 시초가 예상 ({avg_change:+.2f}%)"
    else:
        return f"약보합 시초가 예상 ({avg_change:+.2f}%)"


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 60)
    print("  해외 시장 실시간 모니터링")
    print(f"  현재 KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  활성 시장: {get_active_market()}")
    print("=" * 60)

    # 데이터 수집
    all_markets = []
    grouped = {}

    for category, tickers in MARKETS.items():
        print(f"\n[{category}]")
        grouped[category] = []
        for t_info in tickers:
            data = fetch_market_data(t_info)
            if data:
                all_markets.append(data)
                grouped[category].append(data)
                arrow = "↑" if data["change_pct"] > 0 else "↓" if data["change_pct"] < 0 else "→"
                print(f"  {data['region']} {data['name']:15s} {data['current']:>10.2f} {arrow} {data['change_pct']:+.2f}%")

    # KOSPI 시초가 예측
    kospi_pred = predict_kospi_open(all_markets)

    # 알림 조건 체크
    state = load_state("overseas_market", default={"last_alerts": {}})
    alerts = []
    for market in all_markets:
        need, severity, msg = should_alert(market)
        if need and not is_recent_alert(state, market["ticker"], hours=4):
            alerts.append({"market": market, "severity": severity, "message": msg})
            mark_alert_sent(state, market["ticker"])

    save_state("overseas_market", state)

    print(f"\n[알림 대상] {len(alerts)}건")

    # 텔레그램 전송 (알림이 있는 경우만)
    if alerts:
        msg_lines = [
            "🌎 해외 시장 동향 알림",
            "=" * 25,
            f"\n시각: {datetime.now(KST).strftime('%Y-%m-%d %H:%M KST')}",
            f"활성 시장: {get_active_market()}",
            "",
        ]

        # 카테고리별 알림 정리
        for alert in alerts:
            msg_lines.append(alert["message"])

        # 핵심 지수 요약
        msg_lines.append("\n📊 주요 지수")
        for m in all_markets:
            if m["ticker"] in ["^GSPC", "^IXIC", "^VIX", "^N225"]:
                arrow = "↑" if m["change_pct"] > 0 else "↓"
                msg_lines.append(f"  {m['region']} {m['name']}: {m['current']} {arrow} {m['change_pct']:+.2f}%")

        # KOSPI 시초가 예측
        if kospi_pred and is_kospi_closed():
            msg_lines.append(f"\n📈 KOSPI 시초가: {kospi_pred}")

        msg_lines.append("\n🚨 시뮬레이션. 자동 매매 금지.")
        msg_lines.append("\n대시보드: https://15678910.github.io/ai-finance/")

        ok = send_message("\n".join(msg_lines))
        if ok:
            print(f"  [텔레그램] 알림 전송 완료 ({len(alerts)}건)")
        else:
            print("  [텔레그램] 전송 실패")

    # 결과 저장 (대시보드용)
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "active_market": get_active_market(),
        "kospi_closed": is_kospi_closed(),
        "kospi_open_prediction": kospi_pred,
        "markets": grouped,
        "all_markets": all_markets,
        "alert_count": len(alerts),
        "alerts": [{"name": a["market"]["name"], "severity": a["severity"], "message": a["message"]} for a in alerts],
    }

    # NaN/Infinity → None 변환 (JS JSON.parse 호환)
    import math as _m
    def _clean_nan(obj):
        if isinstance(obj, float):
            return None if (_m.isnan(obj) or _m.isinf(obj)) else obj
        if isinstance(obj, dict):
            return {k: _clean_nan(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [_clean_nan(item) for item in obj]
        return obj
    output = _clean_nan(output)

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, allow_nan=False)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
