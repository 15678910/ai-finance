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
from core.market_hours import drop_incomplete

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
        {"ticker": "^RUT", "name": "Russell 2000", "region": "🇺🇸"},
        {"ticker": "^VIX", "name": "VIX", "region": "🇺🇸", "is_vix": True},
        {"ticker": "^SOX", "name": "필라델피아 반도체", "region": "🇺🇸", "is_semi": True},
    ],
    "미국_선물": [
        {"ticker": "ES=F", "name": "S&P 선물", "region": "🇺🇸", "is_futures": True},
        {"ticker": "NQ=F", "name": "Nasdaq 선물", "region": "🇺🇸", "is_futures": True},
        {"ticker": "YM=F", "name": "Dow 선물", "region": "🇺🇸", "is_futures": True},
    ],
    "아시아": [
        {"ticker": "EWY", "name": "iShares MSCI 한국 ETF", "region": "🇰🇷", "is_korea_proxy": True},
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
        {"ticker": "JPY=X", "name": "USD/JPY", "region": "🇯🇵"},
    ],
    "암호화폐": [
        # BTC는 24시간 거래 — bitcoin_standard.json(1일 2회)만으론 최대 12시간 지연되어
        # 노후 준비 패널의 BTC 수량 목표가 옛 시세로 계산됨. 시간당 도는 이 수집기에 함께 담는다.
        {"ticker": "BTC-USD", "name": "비트코인", "region": "🟠"},
    ],
    "원자재": [
        {"ticker": "GC=F", "name": "금 (Gold)", "region": "🥇"},
        {"ticker": "CL=F", "name": "WTI 원유", "region": "🛢️"},
        {"ticker": "BZ=F", "name": "브렌트 원유", "region": "🛢️"},
        {"ticker": "SI=F", "name": "은 (Silver)", "region": "⚪"},
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
            # 현재가(위 2번)는 실시간이 목적이라 그대로 두고, 52주 고저만 확정 종가 기준으로 계산
            hist_year, _ = drop_incomplete(hist_year, ticker)
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
            "is_semi":   ticker_info.get("is_semi", False),
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


# 시초가 예측 가중치 — 실측으로 결정 (아래 EVIDENCE 참조)
OPEN_W_EWY = 0.2                    # 한국 ETF(EWY) 비중
OPEN_W_FUT = 1 - OPEN_W_EWY         # 미국 선물 비중

# EVIDENCE (2026-08-02 측정) — KOSPI 지수 일별 시가/종가(pykrx) vs 직전 미국 세션
#   표본 481거래일(2024-01~2026-07) · 앞 70% 학습 / 뒤 30%(145일) 검증
#   검증구간 성능 (r · 방향일치 · 시초가갭 MAE):
#     0.6·EWY+0.4·선물 (기존)  r=+0.509  73.8%  1.93%p   ← 기준선(항상 보합) 1.77%p보다 나쁨
#     0.2·EWY+0.8·선물 (채택)  r=+0.557  76.6%  1.44%p
#     선물 단독                r=+0.561  75.9%  1.48%p
#     EWY 단독                r=+0.477  71.7%  2.77%p
#   → 기존 주석("EWY가 선물보다 예측력 높음")은 데이터와 반대였다. 선물 비중을 올려야 맞다.
#   ⚠️ 방향은 4번 중 3번 맞지만 '폭'은 신뢰도가 낮다(MAE 1.44%p vs 갭 표준편차 2.24%p).
#   ⚠️ 시초가 갭 → 그날 장중(시가→종가) 방향 예측력은 없음(방향일치 50.2%, R²=0.03).
#
# 야간선물(코스피200 야간물)을 쓰지 않는 이유:
#   ① pykrx는 선물 OHLCV 미구현(get_future_ohlcv → NotImplementedError) — 티커 목록만 제공.
#      무료로 야간 세션 시세를 받을 경로가 없다.
#   ② 야간선물이 거래되는 시간대는 미국 정규장과 겹친다 = NQ 선물·EWY와 '같은 정보 창구'다.
#      야간선물이 시초가를 움직인다기보다, 둘 다 같은 야간 뉴스를 반영한 결과다.
#      따라서 야간선물을 넣어도 새 정보가 아니라 중복 신호에 가깝다.
def kospi_open_signal(markets: list):
    """합성 신호값(%) 반환 — 채점기(kospi_open_tracker)가 같은 값을 쓰도록 분리."""
    futures = [m for m in markets if m.get("is_futures") and m.get("change_pct") is not None]
    ewy = next((m for m in markets if m.get("is_korea_proxy") and m.get("change_pct") is not None), None)
    fut_change = (sum(m["change_pct"] for m in futures) / len(futures)) if futures else None
    ewy_change = ewy["change_pct"] if ewy else None
    if ewy_change is not None and fut_change is not None:
        return OPEN_W_EWY * ewy_change + OPEN_W_FUT * fut_change, ewy_change, fut_change
    if ewy_change is not None:
        return ewy_change, ewy_change, None
    if fut_change is not None:
        return fut_change, None, fut_change
    return None, None, None


def predict_kospi_open(markets: list) -> str:
    """미국 선물 + EWY(한국 ETF) 기반 KOSPI 시초가 예측 (가중치는 위 EVIDENCE로 실측 결정)."""
    avg_change, _, _ = kospi_open_signal(markets)
    if avg_change is None:
        return ""
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
    # Windows cp949 콘솔에서 이모지 출력 오류 방지
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

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
        # 채점기(kospi_open_tracker.py)가 매일 기록·대조할 수 있도록 숫자 신호도 함께 저장
        "kospi_open_signal": (lambda t: {"composite_pct": None if t[0] is None else round(t[0], 3),
                                         "ewy_pct": None if t[1] is None else round(t[1], 3),
                                         "futures_pct": None if t[2] is None else round(t[2], 3)})(
            kospi_open_signal(all_markets)),
        # 예측 신뢰도 실측치 — 대시보드가 '얼마나 믿을 값인지' 함께 보여주도록 동봉
        "kospi_open_accuracy": {
            "weights": {"ewy": OPEN_W_EWY, "us_futures": round(OPEN_W_FUT, 2)},
            "sample_days": 145, "period": "2024-01~2026-07 중 검증구간(뒤 30%)",
            "direction_hit_pct": 76.6, "corr": 0.557,
            "mae_pp": 1.44, "gap_std_pp": 2.24, "naive_mae_pp": 1.77,
            "intraday_hit_pct": 50.2,
            "note": ("직전 미국 세션(선물 80%·EWY 20%)으로 다음 KRX 개장 갭을 추정. "
                     "방향은 약 4번 중 3번 맞았으나 '폭'의 오차(MAE 1.44%p)는 갭 변동폭(2.24%p) 대비 작지 않다. "
                     "시초가 갭이 그날 장중 방향까지 예측하지는 못한다(방향일치 50.2%). "
                     "코스피200 야간선물은 같은 시간대(미국장)에 거래돼 동일 정보를 반영하므로 별도 추가 이득이 크지 않고, "
                     "무료 데이터 경로도 없어 미사용."),
        },
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
