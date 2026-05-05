"""
비트코인 본위제 / 컴퓨트 달러 모니터링 시스템
=============================================

미국의 비트코인 본위제(Bitcoin Standard) 전환 가능성을 추적하기 위해
관찰 가능한 시장 지표를 수집합니다.

수집 대상:
  1. 비트코인: 가격, 200일 이평선, 변동성, 도미넌스
  2. 스테이블코인 시총: USDT, USDC, DAI (디지털 달러 수요)
  3. 컴퓨트 수요: NVDA, AVGO, AMD (AI 반도체)
  4. 매크로: 달러 인덱스, 미 10Y, 금
  5. 한국 영향: 삼성전자, SK하이닉스, KOSPI

핵심 지표:
  - BTC vs 200MA (강세 확인)
  - 스테이블코인 시총 추이 (디지털 달러 채택)
  - BTC/금 비율 (안전자산 경쟁)
  - 한국 반도체 컴퓨트 노출도

🚨 시뮬레이션/분석용. 정책은 추측 영역.
"""

import os
import sys
import json
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta

try:
    import yfinance as yf
    import pandas as pd
    import numpy as np
except ImportError as e:
    print(f"[오류] 라이브러리 미설치: {e}")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "bitcoin_standard.json")
STATE_FILE = os.path.join(BASE_DIR, "docs", "bitcoin_standard_state.json")
KST = timezone(timedelta(hours=9))


# ====================================================================
# 비트코인 / 컴퓨트 수요 종목
# ====================================================================
TARGETS = {
    "암호화폐": [
        {"ticker": "BTC-USD", "name": "Bitcoin", "key": "btc"},
        {"ticker": "ETH-USD", "name": "Ethereum", "key": "eth"},
    ],
    "AI_컴퓨트": [
        {"ticker": "NVDA", "name": "NVIDIA", "key": "nvda"},
        {"ticker": "AVGO", "name": "Broadcom", "key": "avgo"},
        {"ticker": "AMD", "name": "AMD", "key": "amd"},
        {"ticker": "SOXX", "name": "반도체 ETF", "key": "soxx"},
    ],
    "한국_컴퓨트": [
        {"ticker": "005930.KS", "name": "삼성전자", "key": "samsung"},
        {"ticker": "000660.KS", "name": "SK하이닉스", "key": "skhynix"},
        {"ticker": "042700.KS", "name": "한미반도체", "key": "hanmiSemi"},
    ],
    "매크로": [
        {"ticker": "GC=F", "name": "금", "key": "gold"},
        {"ticker": "DX-Y.NYB", "name": "달러 인덱스", "key": "dxy"},
        {"ticker": "^TNX", "name": "미 10Y 국채", "key": "us10y"},
    ],
}


# ====================================================================
# yfinance 데이터 수집
# ====================================================================
def fetch_market_data(ticker: str, period: str = "1y") -> dict:
    """단일 종목 데이터 수집 + 200MA + 변동성."""
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period=period)
        if hist.empty or len(hist) < 30:
            return None

        current = float(hist["Close"].iloc[-1])
        prev_1d = float(hist["Close"].iloc[-2]) if len(hist) >= 2 else current
        prev_1w = float(hist["Close"].iloc[-6]) if len(hist) >= 6 else current
        prev_1m = float(hist["Close"].iloc[-22]) if len(hist) >= 22 else current
        prev_3m = float(hist["Close"].iloc[-66]) if len(hist) >= 66 else current
        prev_1y = float(hist["Close"].iloc[0]) if len(hist) >= 200 else current

        # 200일 이평선
        ma200 = float(hist["Close"].rolling(200).mean().iloc[-1]) if len(hist) >= 200 else None

        # 30일 변동성
        returns_30d = hist["Close"].pct_change().iloc[-30:]
        volatility_30d = float(returns_30d.std() * np.sqrt(252) * 100) if len(returns_30d) > 5 else None

        return {
            "ticker": ticker,
            "current": round(current, 2),
            "change_1d": round((current - prev_1d) / prev_1d * 100, 2) if prev_1d > 0 else 0,
            "change_1w": round((current - prev_1w) / prev_1w * 100, 2) if prev_1w > 0 else 0,
            "change_1m": round((current - prev_1m) / prev_1m * 100, 2) if prev_1m > 0 else 0,
            "change_3m": round((current - prev_3m) / prev_3m * 100, 2) if prev_3m > 0 else 0,
            "change_1y": round((current - prev_1y) / prev_1y * 100, 2) if prev_1y > 0 else 0,
            "ma200": round(ma200, 2) if ma200 else None,
            "above_ma200": (current > ma200) if ma200 else None,
            "volatility_30d_pct": round(volatility_30d, 2) if volatility_30d else None,
        }
    except Exception as e:
        print(f"  [실패] {ticker}: {e}")
        return None


# ====================================================================
# CoinGecko 스테이블코인 시총 수집
# ====================================================================
def fetch_stablecoin_caps() -> dict:
    """주요 스테이블코인 시가총액 수집."""
    print("\n[CoinGecko] 스테이블코인 시총 수집...")
    coin_ids = "tether,usd-coin,dai,first-digital-usd,paypal-usd"
    url = f"https://api.coingecko.com/api/v3/simple/price?ids={coin_ids}&vs_currencies=usd&include_market_cap=true&include_24hr_change=true"

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        result = {}
        name_map = {
            "tether": "USDT (Tether)",
            "usd-coin": "USDC (Circle)",
            "dai": "DAI",
            "first-digital-usd": "FDUSD",
            "paypal-usd": "PYUSD",
        }
        total_cap = 0
        for coin_id, info in data.items():
            cap = info.get("usd_market_cap", 0)
            change = info.get("usd_24h_change", 0)
            result[coin_id] = {
                "name": name_map.get(coin_id, coin_id),
                "market_cap_usd": int(cap),
                "market_cap_billion": round(cap / 1e9, 2),
                "change_24h_pct": round(change, 2) if change else 0,
            }
            total_cap += cap

        result["__total__"] = {
            "name": "총 스테이블코인 시총",
            "market_cap_usd": int(total_cap),
            "market_cap_billion": round(total_cap / 1e9, 2),
        }
        print(f"  총 시총: ${total_cap/1e9:.2f}B")
        return result
    except Exception as e:
        print(f"  [실패] CoinGecko: {e}")
        return {}


def fetch_btc_dominance() -> dict:
    """비트코인 시장 도미넌스."""
    print("[CoinGecko] BTC 도미넌스 수집...")
    url = "https://api.coingecko.com/api/v3/global"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        market_data = data.get("data", {})
        market_cap_pct = market_data.get("market_cap_percentage", {})

        btc_dominance = market_cap_pct.get("btc", 0)
        eth_dominance = market_cap_pct.get("eth", 0)
        total_market_cap_usd = market_data.get("total_market_cap", {}).get("usd", 0)

        return {
            "btc_dominance_pct": round(btc_dominance, 2),
            "eth_dominance_pct": round(eth_dominance, 2),
            "total_market_cap_billion": round(total_market_cap_usd / 1e9, 2),
        }
    except Exception as e:
        print(f"  [실패] BTC dominance: {e}")
        return {}


# ====================================================================
# 비트코인 본위제 점수 계산
# ====================================================================
def calculate_bts_score(btc: dict, stablecoins: dict, dominance: dict, semis: dict) -> dict:
    """비트코인 본위제 추진 가능성 점수 (0-100).
    높을수록 시장이 본위제 시나리오를 반영하고 있음을 의미."""
    score = 0
    factors = []

    # 1) BTC 강세 (200MA 위, 1년 +50% 이상)
    if btc.get("above_ma200"):
        score += 15
        factors.append("BTC 200MA 상회 (강세 추세)")
    btc_1y = btc.get("change_1y", 0)
    if btc_1y > 100:
        score += 25
        factors.append(f"BTC 1년 +{btc_1y:.0f}% (대규모 상승)")
    elif btc_1y > 50:
        score += 15
        factors.append(f"BTC 1년 +{btc_1y:.0f}% (강한 상승)")
    elif btc_1y > 20:
        score += 8
        factors.append(f"BTC 1년 +{btc_1y:.0f}%")

    # 2) 스테이블코인 시총 폭증 (디지털 달러 채택)
    total_cap_b = stablecoins.get("__total__", {}).get("market_cap_billion", 0)
    if total_cap_b > 500:
        score += 20
        factors.append(f"스테이블코인 시총 ${total_cap_b:.0f}B 돌파")
    elif total_cap_b > 300:
        score += 12
        factors.append(f"스테이블코인 시총 ${total_cap_b:.0f}B (성장세)")
    elif total_cap_b > 200:
        score += 6

    # 3) BTC 도미넌스 (전통 자산 vs 알트 약세 = BTC 본위제 시나리오 강화)
    btc_dom = dominance.get("btc_dominance_pct", 0)
    if btc_dom > 60:
        score += 15
        factors.append(f"BTC 도미넌스 {btc_dom:.1f}% (본위제 시나리오 강화)")
    elif btc_dom > 55:
        score += 10
        factors.append(f"BTC 도미넌스 {btc_dom:.1f}%")

    # 4) AI 컴퓨트 수요 (NVDA, SOXX 강세)
    nvda_3m = semis.get("nvda", {}).get("change_3m", 0)
    if nvda_3m > 30:
        score += 15
        factors.append(f"NVDA 3M +{nvda_3m:.0f}% (AI 컴퓨트 수요 폭발)")
    elif nvda_3m > 10:
        score += 8

    # 5) BTC 변동성 안정화 (성숙도 = 본위제 적합성)
    btc_vol = btc.get("volatility_30d_pct", 0)
    if btc_vol and btc_vol < 50:
        score += 10
        factors.append(f"BTC 변동성 {btc_vol:.0f}% (성숙도 진전)")

    score = min(100, max(0, score))

    if score >= 70:
        level = "본위제 시나리오 강력"
        severity = "red"
    elif score >= 50:
        level = "본위제 신호 부상"
        severity = "amber"
    elif score >= 30:
        level = "관찰 필요"
        severity = "amber"
    else:
        level = "약한 신호"
        severity = "green"

    return {
        "score": score,
        "level": level,
        "severity": severity,
        "factors": factors,
    }


# ====================================================================
# 한국 영향 분석
# ====================================================================
def analyze_korea_impact(score_data: dict, kor_semis: dict, btc: dict) -> list:
    """한국 시장에 미치는 영향 분석."""
    impacts = []
    score = score_data["score"]

    if score >= 60:
        impacts.append({
            "category": "한국 반도체",
            "impact": "긍정",
            "description": "컴퓨트 달러 시나리오 → 미국 AI 인프라 투자 강제 → 삼성전자/SK하이닉스 HBM 수요 폭증 가능. 비트코인 본위제는 반도체 수퍼사이클 트리거.",
        })
        impacts.append({
            "category": "원화 환율",
            "impact": "복합",
            "description": "스테이블코인이 달러 수요를 흡수 → 달러 강세 → 원화 약세 압력. 단, 한국 수출주에는 유리.",
        })
        impacts.append({
            "category": "한국 디지털 자산 정책",
            "impact": "긴급",
            "description": "한국이 비트코인 미보유 → 미국 디지털 패권 체제에서 후순위. 정부 차원의 비트코인 비축 검토 시급.",
        })

    if score >= 50:
        # NVDA 상승 시 한국 반도체 동조
        nvda_3m = kor_semis.get("nvda", {}).get("change_3m", 0) if isinstance(kor_semis.get("nvda"), dict) else 0
        sam_3m = kor_semis.get("samsung", {}).get("change_3m", 0)
        if nvda_3m > 20 and sam_3m < nvda_3m / 2:
            impacts.append({
                "category": "한국 반도체 갭 (디스카운트)",
                "impact": "긍정",
                "description": f"NVDA 3M +{nvda_3m:.0f}% vs 삼성전자 +{sam_3m:.0f}% - 한국 반도체 follow 매수 기회 가능.",
            })

    # BTC 200MA 돌파 시
    if btc.get("above_ma200") and btc.get("change_1m", 0) > 5:
        impacts.append({
            "category": "한국 암호화폐 거래소",
            "impact": "긍정",
            "description": "BTC 200MA 상회 + 모멘텀 → 두나무, 빗썸 등 거래량 증가. 김치 프리미엄 모니터링.",
        })

    return impacts


# ====================================================================
# 텔레그램 알림
# ====================================================================
def send_telegram(score_data: dict, btc: dict, stablecoins: dict, alerts: list):
    if not alerts:
        return

    env_path = os.path.join(CONFIG_DIR, ".env")
    bot_token, chat_id = None, None
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                if "=" not in line or line.startswith("#"):
                    continue
                k, v = line.strip().split("=", 1)
                v = v.strip().strip("'\"")
                if k.strip() == "TELEGRAM_FINANCE_BOT_TOKEN":
                    bot_token = v
                elif k.strip() == "TELEGRAM_FINANCE_CHAT_ID":
                    chat_id = v
    bot_token = bot_token or os.environ.get("TELEGRAM_FINANCE_BOT_TOKEN")
    chat_id = chat_id or os.environ.get("TELEGRAM_FINANCE_CHAT_ID")

    if not bot_token or not chat_id:
        return

    lines = [
        "🪙 비트코인 본위제 모니터",
        "=" * 25,
        "",
        f"점수: {score_data['score']}/100 ({score_data['level']})",
        "",
        f"BTC: ${btc.get('current', 0):,.0f} ({btc.get('change_1d', 0):+.1f}% 1D, {btc.get('change_1y', 0):+.0f}% 1Y)",
        f"스테이블코인 시총: ${stablecoins.get('__total__', {}).get('market_cap_billion', 0):.0f}B",
        "",
        "주요 알림:",
    ]
    for a in alerts:
        lines.append(f"  • {a}")

    lines.append("\n🚨 시뮬레이션. 정책은 추측 영역.")
    lines.append("\n대시보드: https://15678910.github.io/ai-finance/")

    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        body = urllib.parse.urlencode({"chat_id": chat_id, "text": "\n".join(lines)}).encode()
        req = urllib.request.Request(url, data=body, method="POST")
        with urllib.request.urlopen(req, timeout=10) as resp:
            json.loads(resp.read())
        print("  [텔레그램] 전송 완료")
    except Exception as e:
        print(f"  [텔레그램] 실패: {e}")


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


def check_alerts(score_data, btc, stablecoins, state) -> list:
    alerts = []
    now = datetime.now(timezone.utc).isoformat()

    def is_recent(key, hours=24):
        last = state.get("last_alerts", {}).get(key)
        if not last:
            return False
        try:
            last_dt = datetime.fromisoformat(last)
            return (datetime.now(timezone.utc) - last_dt) < timedelta(hours=hours)
        except Exception:
            return False

    def mark(key):
        state.setdefault("last_alerts", {})[key] = now

    # 점수 70+ 알림
    if score_data["score"] >= 70 and not is_recent("score_70"):
        alerts.append(f"🔴 본위제 점수 {score_data['score']}/100 - {score_data['level']}")
        mark("score_70")

    # BTC 신고가 돌파 (1D +10% 이상)
    btc_1d = btc.get("change_1d", 0)
    if btc_1d > 10 and not is_recent("btc_surge"):
        alerts.append(f"🔴 BTC 1일 +{btc_1d:.1f}% 급등 (현재가 ${btc.get('current', 0):,.0f})")
        mark("btc_surge")
    elif btc_1d < -10 and not is_recent("btc_dump"):
        alerts.append(f"🔴 BTC 1일 {btc_1d:.1f}% 급락")
        mark("btc_dump")

    # 스테이블코인 시총 마일스톤
    total_b = stablecoins.get("__total__", {}).get("market_cap_billion", 0)
    if total_b > 500 and not is_recent("stable_500"):
        alerts.append(f"🟠 스테이블코인 시총 ${total_b:.0f}B 돌파 (디지털 달러 채택 가속)")
        mark("stable_500")

    return alerts


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 65)
    print("  비트코인 본위제 / 컴퓨트 달러 모니터")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    # 시장 데이터 수집
    grouped = {}
    flat = {}
    for category, items in TARGETS.items():
        print(f"\n[{category}]")
        grouped[category] = []
        for t in items:
            d = fetch_market_data(t["ticker"])
            if d:
                d["name"] = t["name"]
                d["key"] = t["key"]
                grouped[category].append(d)
                flat[t["key"]] = d
                arrow = "↑" if d["change_1d"] > 0 else "↓" if d["change_1d"] < 0 else "→"
                ma_status = "✓" if d.get("above_ma200") else "✗" if d.get("above_ma200") is False else "?"
                print(f"  {t['name']:18s} {d['current']:>12,.2f} {arrow} {d['change_1d']:+5.2f}% (1Y {d['change_1y']:+6.1f}%) MA200 {ma_status}")

    # 스테이블코인 / 도미넌스
    stablecoins = fetch_stablecoin_caps()
    dominance = fetch_btc_dominance()

    # 점수 계산
    btc_data = flat.get("btc", {})
    score_data = calculate_bts_score(btc_data, stablecoins, dominance, flat)
    print(f"\n[비트코인 본위제 점수] {score_data['score']}/100 ({score_data['level']})")
    for f in score_data["factors"]:
        print(f"  • {f}")

    # 한국 영향
    impacts = analyze_korea_impact(score_data, flat, btc_data)
    if impacts:
        print(f"\n[한국 영향 분석] {len(impacts)}건")
        for imp in impacts:
            print(f"  [{imp['impact']}] {imp['category']}")

    # 알림
    state = load_state()
    alerts = check_alerts(score_data, btc_data, stablecoins, state)
    save_state(state)
    if alerts:
        send_telegram(score_data, btc_data, stablecoins, alerts)

    # 결과 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "bts_score": score_data,
        "btc": btc_data,
        "markets": grouped,
        "stablecoins": stablecoins,
        "dominance": dominance,
        "korea_impacts": impacts,
        "alerts": alerts,
        "warning": "🚨 시뮬레이션. 정책 시점/내용은 추측 영역.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    print("\n" + "=" * 65)
    print(f"  점수: {score_data['score']}/100 ({score_data['level']})")
    print(f"  BTC: ${btc_data.get('current', 0):,.0f} ({btc_data.get('change_1y', 0):+.1f}% 1Y)")
    print(f"  스테이블코인 시총: ${stablecoins.get('__total__', {}).get('market_cap_billion', 0):.0f}B")
    print("  ⚠️ 시뮬레이션 / 정책 추측 영역.")
    print("=" * 65)


if __name__ == "__main__":
    main()
