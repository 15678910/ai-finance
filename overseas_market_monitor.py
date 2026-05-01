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

try:
    import yfinance as yf
except ImportError:
    print("[오류] yfinance 미설치. pip install yfinance")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "overseas_market.json")
STATE_FILE = os.path.join(BASE_DIR, "docs", "overseas_market_state.json")

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
    """단일 시장 데이터 수집."""
    ticker = ticker_info["ticker"]
    try:
        stock = yf.Ticker(ticker)
        # 최근 5일 데이터로 변동 계산
        hist = stock.history(period="5d")
        if hist.empty or len(hist) < 2:
            return None

        current = float(hist["Close"].iloc[-1])
        prev = float(hist["Close"].iloc[-2])
        change_pct = (current - prev) / prev * 100 if prev > 0 else 0

        # 52주 데이터
        hist_year = stock.history(period="1y")
        high_52w = float(hist_year["High"].max()) if not hist_year.empty else None
        low_52w = float(hist_year["Low"].min()) if not hist_year.empty else None

        return {
            "ticker": ticker,
            "name": ticker_info["name"],
            "region": ticker_info.get("region", ""),
            "current": round(current, 2),
            "previous": round(prev, 2),
            "change_pct": round(change_pct, 2),
            "high_52w": round(high_52w, 2) if high_52w else None,
            "low_52w": round(low_52w, 2) if low_52w else None,
            "is_vix": ticker_info.get("is_vix", False),
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
# 텔레그램
# ====================================================================
def parse_env(env_path: str) -> dict:
    env_vars = {}
    if not os.path.exists(env_path):
        return env_vars
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k = k.strip()
            v = v.strip().strip("'\"")
            env_vars[k] = v
    return env_vars


def telegram_send(bot_token, chat_id, text):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    data = urllib.parse.urlencode({
        "chat_id": chat_id,
        "text": text,
        "disable_web_page_preview": "true",
    }).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    with urllib.request.urlopen(req, timeout=10) as resp:
        result = json.loads(resp.read().decode("utf-8"))
        if not result.get("ok"):
            raise RuntimeError(f"Telegram error: {result}")


# ====================================================================
# 상태 관리 (중복 알림 방지)
# ====================================================================
def load_state():
    if not os.path.exists(STATE_FILE):
        return {"last_alerts": {}}
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"last_alerts": {}}


def save_state(state):
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def is_alert_recent(state, ticker, hours=4):
    """최근 N시간 내 알림 발송 여부."""
    last = state.get("last_alerts", {}).get(ticker)
    if not last:
        return False
    try:
        last_dt = datetime.fromisoformat(last)
        now = datetime.now(timezone.utc)
        return (now - last_dt) < timedelta(hours=hours)
    except Exception:
        return False


def mark_alert_sent(state, ticker):
    state.setdefault("last_alerts", {})[ticker] = datetime.now(timezone.utc).isoformat()


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
    state = load_state()
    alerts = []
    for market in all_markets:
        need, severity, msg = should_alert(market)
        if need and not is_alert_recent(state, market["ticker"]):
            alerts.append({"market": market, "severity": severity, "message": msg})
            mark_alert_sent(state, market["ticker"])

    save_state(state)

    print(f"\n[알림 대상] {len(alerts)}건")

    # 텔레그램 전송 (알림이 있는 경우만)
    env_vars = parse_env(os.path.join(CONFIG_DIR, ".env"))
    bot_token = env_vars.get("TELEGRAM_FINANCE_BOT_TOKEN") or os.environ.get("TELEGRAM_FINANCE_BOT_TOKEN")
    chat_id = env_vars.get("TELEGRAM_FINANCE_CHAT_ID") or os.environ.get("TELEGRAM_FINANCE_CHAT_ID")

    if alerts and bot_token and chat_id:
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

        try:
            telegram_send(bot_token, chat_id, "\n".join(msg_lines))
            print(f"  [텔레그램] 알림 전송 완료 ({len(alerts)}건)")
        except Exception as e:
            print(f"  [텔레그램] 전송 실패: {e}")

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

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
