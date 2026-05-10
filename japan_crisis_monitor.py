"""
일본 경제 위기 모니터링 시스템
================================

엔화 캐리 트레이드 청산 / 일본 금리 인상 / 인구 위기로 인한
글로벌 금융 시스템 충격을 추적하고 조기 경고를 발송합니다.

수집 대상:
  - 시장 (yfinance): USD/JPY, Nikkei, TOPIX
  - 매크로 (FRED): 일본 10Y 금리, 정책금리, CPI, GDP대비 부채, 인구

핵심 지표:
  - 엔화 캐리 트레이드 압력 지수 (0-100)
  - 일미 금리차 (캐리 매력도)
  - 위기 종합 점수 (0-100)

알림 조건:
  🔴 긴급: USD/JPY 160+ 돌파, 엔화 1주 +5% 강세 (캐리 청산)
  🟠 경고: 캐리 압력 지수 70+, 일본 CPI 3%+
  🟡 주의: 위기 점수 50+

🚨 시뮬레이션/분석용. 자동 매매 절대 금지.
"""

import os
import sys
import json
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta
from io import StringIO

from core import send_message, load_state, save_state, is_recent_alert, mark_alert_sent

try:
    import yfinance as yf
    import pandas as pd
    import numpy as np
except ImportError as e:
    print(f"[오류] 라이브러리 미설치: {e}")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "japan_crisis.json")

KST = timezone(timedelta(hours=9))


# ====================================================================
# FRED 데이터 수집 (API 키 불필요, CSV 다운로드)
# ====================================================================
def fetch_fred_csv(series_id: str, start_date: str = None, retries: int = 3) -> pd.DataFrame:
    """FRED CSV로부터 시계열 데이터 수집 (재시도 포함)."""
    if not start_date:
        start_date = (datetime.now() - timedelta(days=730)).strftime("%Y-%m-%d")

    url = "https://fred.stlouisfed.org/graph/fredgraph.csv"
    params = {
        "id": series_id,
        "cosd": start_date,
        "coed": datetime.now().strftime("%Y-%m-%d"),
    }
    full_url = f"{url}?{urllib.parse.urlencode(params)}"

    import time as _time
    last_err = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(full_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                content = resp.read().decode("utf-8")

            df = pd.read_csv(StringIO(content))
            df.columns = ["DATE", "VALUE"]
            df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
            df["VALUE"] = pd.to_numeric(df["VALUE"], errors="coerce")
            df = df.dropna()
            return df
        except Exception as e:
            last_err = e
            if attempt < retries - 1:
                _time.sleep(2 ** attempt)  # 지수적 백오프

    print(f"  [FRED 실패] {series_id}: {last_err}")
    return pd.DataFrame()


def get_latest_fred(series_id: str, label: str) -> dict:
    """FRED 시리즈의 최신값과 변화량."""
    df = fetch_fred_csv(series_id)
    if df.empty:
        return {"label": label, "current": None, "change_3m": None, "change_1y": None}

    current = float(df["VALUE"].iloc[-1])
    current_date = df["DATE"].iloc[-1]

    # 3개월 전 값
    three_months_ago = current_date - timedelta(days=90)
    df_3m = df[df["DATE"] <= three_months_ago]
    val_3m = float(df_3m["VALUE"].iloc[-1]) if not df_3m.empty else None

    # 1년 전 값
    one_year_ago = current_date - timedelta(days=365)
    df_1y = df[df["DATE"] <= one_year_ago]
    val_1y = float(df_1y["VALUE"].iloc[-1]) if not df_1y.empty else None

    return {
        "label": label,
        "current": round(current, 3),
        "current_date": current_date.strftime("%Y-%m-%d"),
        "previous_3m": round(val_3m, 3) if val_3m else None,
        "previous_1y": round(val_1y, 3) if val_1y else None,
        "change_3m": round(current - val_3m, 3) if val_3m else None,
        "change_1y": round(current - val_1y, 3) if val_1y else None,
    }


# ====================================================================
# yfinance 데이터 수집
# ====================================================================
def fetch_yf_market(ticker: str, label: str, period: str = "3mo") -> dict:
    """yfinance에서 시장 데이터 수집."""
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period=period)
        if hist.empty:
            return {"label": label, "current": None}

        current = float(hist["Close"].iloc[-1])
        prev_1d = float(hist["Close"].iloc[-2]) if len(hist) >= 2 else current

        # 1주 전 (5거래일)
        prev_1w = float(hist["Close"].iloc[-6]) if len(hist) >= 6 else current
        # 1개월 전 (~22거래일)
        prev_1m = float(hist["Close"].iloc[-22]) if len(hist) >= 22 else current

        change_1d = (current - prev_1d) / prev_1d * 100 if prev_1d > 0 else 0
        change_1w = (current - prev_1w) / prev_1w * 100 if prev_1w > 0 else 0
        change_1m = (current - prev_1m) / prev_1m * 100 if prev_1m > 0 else 0

        return {
            "label": label,
            "ticker": ticker,
            "current": round(current, 2),
            "change_1d": round(change_1d, 2),
            "change_1w": round(change_1w, 2),
            "change_1m": round(change_1m, 2),
        }
    except Exception as e:
        print(f"  [yfinance 실패] {ticker}: {e}")
        return {"label": label, "current": None}


# ====================================================================
# 캐리 트레이드 압력 지수
# ====================================================================
def calculate_carry_pressure(usd_jpy: dict, jp_10y: dict, us_10y: dict) -> dict:
    """캐리 트레이드 청산 압력 지수 (0-100).

    높을수록 캐리 청산 압력 강함:
    - 엔화 급강세 (USD/JPY 하락) → 청산 신호
    - 일미 금리차 축소 → 캐리 매력도 감소
    """
    score = 0
    factors = []

    # 1) 엔화 1주 변동률 (강세 = 청산 신호)
    yen_w = usd_jpy.get("change_1w") or 0
    if yen_w < -5:  # 엔화 5%+ 강세
        score += 40
        factors.append(f"엔화 1주 +{abs(yen_w):.1f}% 강세 (캐리 청산 가속)")
    elif yen_w < -3:
        score += 25
        factors.append(f"엔화 1주 +{abs(yen_w):.1f}% 강세 (청산 신호)")
    elif yen_w < -1:
        score += 10
        factors.append(f"엔화 약강세 ({yen_w:+.1f}%)")
    elif yen_w > 3:
        factors.append(f"엔화 약세 지속 ({yen_w:+.1f}%) - 캐리 매력 유지")

    # 2) 일미 금리차 축소 (캐리 매력도 감소)
    jp_rate = jp_10y.get("current") or 0
    us_rate = us_10y.get("current") or 0
    spread = us_rate - jp_rate

    if spread < 2.0:
        score += 30
        factors.append(f"일미 금리차 {spread:.2f}%p (캐리 매력 급감)")
    elif spread < 3.0:
        score += 15
        factors.append(f"일미 금리차 {spread:.2f}%p (캐리 매력 약화)")

    # 3) 일본 10Y 금리 상승 추세
    jp_change_3m = jp_10y.get("change_3m") or 0
    if jp_change_3m > 0.3:
        score += 20
        factors.append(f"일본 10Y 3개월 +{jp_change_3m:.2f}%p (BOJ 압박 가속)")
    elif jp_change_3m > 0.1:
        score += 10
        factors.append(f"일본 10Y 상승 ({jp_change_3m:+.2f}%p)")

    # 4) USD/JPY 절대 수준
    usd_jpy_val = usd_jpy.get("current") or 0
    if usd_jpy_val > 160:
        score += 10
        factors.append(f"USD/JPY {usd_jpy_val} (160 돌파, BOJ 개입 압력)")

    score = min(100, max(0, score))

    if score >= 70:
        level = "위험"
        severity = "red"
    elif score >= 50:
        level = "경고"
        severity = "amber"
    elif score >= 30:
        level = "주의"
        severity = "amber"
    else:
        level = "안정"
        severity = "green"

    return {
        "score": score,
        "level": level,
        "severity": severity,
        "spread_us_jp": round(spread, 2),
        "factors": factors,
    }


# ====================================================================
# 종합 위기 점수
# ====================================================================
def calculate_crisis_score(data: dict) -> dict:
    """일본 위기 종합 점수 (0-100)."""
    score = 0
    factors = []

    usd_jpy = data["market"]["usd_jpy"]["current"] or 0
    nikkei_w = data["market"]["nikkei"]["change_1w"] or 0
    cpi = data["macro"]["cpi"]["current"] or 0
    debt_gdp = data["macro"]["debt_gdp"]["current"] or 0
    carry_score = data["carry_pressure"]["score"]

    # USD/JPY (1점=1점, 160=10, 170=20, 180=30)
    if usd_jpy > 180:
        score += 30
        factors.append(f"USD/JPY {usd_jpy} 극심한 엔화 약세")
    elif usd_jpy > 170:
        score += 20
        factors.append(f"USD/JPY {usd_jpy} 심각한 엔화 약세")
    elif usd_jpy > 160:
        score += 10
        factors.append(f"USD/JPY {usd_jpy} 엔화 약세 지속")

    # 캐리 압력 (40% 반영)
    score += carry_score * 0.4
    if carry_score >= 50:
        factors.append(f"캐리 청산 압력 {carry_score}/100")

    # CPI (인플레이션)
    if cpi and cpi > 130:  # CPI Index 기준 (2020=100)
        score += 10
        factors.append(f"일본 CPI {cpi:.1f} (인플레이션 압력)")

    # 부채/GDP
    if debt_gdp and debt_gdp > 250:
        score += 15
        factors.append(f"부채/GDP {debt_gdp:.0f}% (재정 한계)")
    elif debt_gdp and debt_gdp > 200:
        score += 10

    # Nikkei 1주 폭락
    if nikkei_w < -5:
        score += 15
        factors.append(f"Nikkei 1주 {nikkei_w:.1f}% 폭락")

    score = min(100, max(0, score))

    if score >= 70:
        level = "심각"
        severity = "red"
    elif score >= 50:
        level = "경고"
        severity = "amber"
    elif score >= 30:
        level = "주의"
        severity = "amber"
    else:
        level = "안정"
        severity = "green"

    return {
        "score": round(score, 1),
        "level": level,
        "severity": severity,
        "factors": factors,
    }


# ====================================================================
# 글로벌 영향 분석
# ====================================================================
def analyze_global_impact(carry: dict, crisis: dict, data: dict) -> list:
    """일본 위기가 글로벌 시장에 미치는 영향 분석."""
    impacts = []
    carry_score = carry["score"]
    yen_1w = data["market"]["usd_jpy"]["change_1w"] or 0

    # 캐리 트레이드 청산 영향
    if carry_score >= 70:
        impacts.append({
            "category": "미국 기술주",
            "impact": "부정",
            "description": "엔화 부채 상환 위해 AAPL, NVDA, MSFT 등 매도 압력. 단기 변동성 확대.",
        })
        impacts.append({
            "category": "미 국채",
            "impact": "부정",
            "description": "일본 기관의 미 국채 매각 가속 → 미국 10Y 금리 상승 → 주택 대출/신용카드 금리 상승.",
        })
        impacts.append({
            "category": "달러",
            "impact": "부정",
            "description": "엔화 강세 → 달러 약세 압력. 달러 인덱스 하락 가능성.",
        })
        impacts.append({
            "category": "한국 시장 (KOSPI)",
            "impact": "복합",
            "description": "글로벌 캐리 청산으로 위험자산 회피 → 단기 KOSPI 부정. 단, 엔화 강세 시 일본 수출 경쟁력 약화 → 한국 반도체/자동차 반사이익 가능.",
        })
    elif carry_score >= 50:
        impacts.append({
            "category": "글로벌 위험자산",
            "impact": "주의",
            "description": "캐리 청산 우려로 변동성 확대. 신중한 포지션 관리 권고.",
        })

    # 엔화 급강세 시
    if yen_1w < -3:
        impacts.append({
            "category": "수출주 (한국)",
            "impact": "긍정",
            "description": "엔화 강세로 일본 수출 경쟁력 약화 → 한국 수출주 (자동차, 반도체) 반사이익 기대.",
        })

    # 위기 점수 높을 때
    if crisis["score"] >= 70:
        impacts.append({
            "category": "안전자산 (금)",
            "impact": "긍정",
            "description": "글로벌 금융 시스템 위기 가능성 → 금 가격 상승 압력.",
        })

    return impacts


# ====================================================================
# 알림 트리거
# ====================================================================
def check_alerts(data: dict, state: dict) -> list:
    """알림 조건 체크."""
    alerts = []

    usd_jpy = data["market"]["usd_jpy"]
    yen_1w = usd_jpy.get("change_1w") or 0
    yen_now = usd_jpy.get("current") or 0
    carry = data["carry_pressure"]
    crisis = data["crisis"]

    # USD/JPY 160 돌파
    if yen_now > 160 and not is_recent_alert(state, "usd_jpy_160", hours=12):
        alerts.append({
            "severity": "긴급",
            "title": "USD/JPY 160 돌파",
            "message": f"🔴 USD/JPY {yen_now} - 엔화 심각한 약세. BOJ 개입/금리 인상 압력 극대화.",
        })
        mark_alert_sent(state, "usd_jpy_160")

    # 엔화 1주 5% 이상 강세 (캐리 청산)
    if yen_1w < -5 and not is_recent_alert(state, "yen_strong", hours=12):
        alerts.append({
            "severity": "긴급",
            "title": "엔화 급강세 - 캐리 청산 신호",
            "message": f"🔴 엔화 1주 +{abs(yen_1w):.1f}% 강세. 글로벌 자산 매도 압력 증가 (미국 기술주, 국채).",
        })
        mark_alert_sent(state, "yen_strong")

    # 캐리 압력 70+
    if carry["score"] >= 70 and not is_recent_alert(state, "carry_70", hours=12):
        alerts.append({
            "severity": "경고",
            "title": "캐리 트레이드 청산 위험",
            "message": f"🟠 캐리 압력 지수 {carry['score']}/100 ({carry['level']}). 글로벌 변동성 확대 주의.",
        })
        mark_alert_sent(state, "carry_70")

    # 위기 점수 70+
    if crisis["score"] >= 70 and not is_recent_alert(state, "crisis_70", hours=12):
        alerts.append({
            "severity": "긴급",
            "title": "일본 위기 심각 단계",
            "message": f"🔴 위기 종합 점수 {crisis['score']}/100 ({crisis['level']}). 글로벌 금융 충격 가능성.",
        })
        mark_alert_sent(state, "crisis_70")

    return alerts


# ====================================================================
# 텔레그램 전송
# ====================================================================
def send_telegram(data: dict, alerts: list):
    """위기 알림 또는 일일 요약 전송."""
    # 알림이 있을 때만 전송 (일일 요약은 daily 분석에서 처리)
    if not alerts:
        return

    lines = ["🇯🇵 일본 경제 위기 알림", "=" * 25, ""]
    for a in alerts:
        lines.append(f"[{a['severity']}] {a['title']}")
        lines.append(a["message"])
        lines.append("")

    # 핵심 지표 요약
    lines.append("📊 핵심 지표")
    usd_jpy = data["market"]["usd_jpy"]
    nikkei = data["market"]["nikkei"]
    carry = data["carry_pressure"]
    lines.append(f"  USD/JPY: {usd_jpy.get('current')} ({usd_jpy.get('change_1w', 0):+.1f}% 1W)")
    lines.append(f"  Nikkei: {nikkei.get('current')} ({nikkei.get('change_1w', 0):+.1f}% 1W)")
    lines.append(f"  캐리 압력: {carry['score']}/100 ({carry['level']})")
    lines.append(f"  위기 점수: {data['crisis']['score']}/100 ({data['crisis']['level']})")

    lines.append("\n🚨 시뮬레이션. 자동 매매 금지.")
    lines.append("\n대시보드: https://15678910.github.io/ai-finance/")

    ok = send_message("\n".join(lines))
    if ok:
        print(f"  [텔레그램] 전송 완료 ({len(alerts)}건 알림)")
    else:
        print("  [텔레그램] 전송 실패")


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 65)
    print("  일본 경제 위기 모니터링")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    # 시장 데이터 수집
    print("\n[시장 데이터]")
    market = {
        "usd_jpy": fetch_yf_market("JPY=X", "USD/JPY"),
        "nikkei": fetch_yf_market("^N225", "Nikkei 225"),
        "topix": fetch_yf_market("1306.T", "TOPIX (ETF)"),  # TOPIX ETF
        "wti": fetch_yf_market("CL=F", "WTI 유가"),
    }
    for k, v in market.items():
        if v.get("current"):
            print(f"  {v['label']:15s}: {v['current']} ({v.get('change_1w', 0):+.2f}% 1W, {v.get('change_1m', 0):+.2f}% 1M)")

    # FRED 매크로 데이터
    print("\n[FRED 매크로 데이터]")
    macro = {
        "policy_rate": get_latest_fred("INTDSRJPM193N", "BOJ 정책금리"),
        "yield_10y": get_latest_fred("IRLTLT01JPM156N", "일본 10Y 국채"),
        "us_10y": get_latest_fred("DGS10", "미국 10Y 국채"),
        "cpi": get_latest_fred("JPNCPIALLMINMEI", "일본 CPI"),
        "debt_gdp": get_latest_fred("GGGDTAJPA188N", "GDP 대비 부채(%)"),
        "population": get_latest_fred("POPTOTJPA647NWDB", "일본 인구"),
    }
    for k, v in macro.items():
        if v["current"] is not None:
            print(f"  {v['label']:20s}: {v['current']} (3M: {v.get('change_3m')})")

    # 캐리 트레이드 압력
    carry = calculate_carry_pressure(market["usd_jpy"], macro["yield_10y"], macro["us_10y"])
    print(f"\n[캐리 트레이드 압력] {carry['score']}/100 ({carry['level']})")
    for f in carry["factors"]:
        print(f"  • {f}")

    # 데이터 통합
    data = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "market": market,
        "macro": macro,
        "carry_pressure": carry,
    }

    # 위기 종합 점수
    crisis = calculate_crisis_score(data)
    data["crisis"] = crisis
    print(f"\n[위기 종합 점수] {crisis['score']}/100 ({crisis['level']})")
    for f in crisis["factors"]:
        print(f"  • {f}")

    # 글로벌 영향 분석
    impacts = analyze_global_impact(carry, crisis, data)
    data["global_impacts"] = impacts
    if impacts:
        print(f"\n[글로벌 영향 분석] {len(impacts)}건")
        for imp in impacts:
            print(f"  [{imp['impact']}] {imp['category']}: {imp['description'][:80]}")

    # 알림 체크
    state = load_state("japan_crisis", default={"last_alerts": {}})
    alerts = check_alerts(data, state)
    save_state("japan_crisis", state)
    data["alerts"] = alerts

    # 텔레그램 전송
    if alerts:
        send_telegram(data, alerts)

    # 결과 저장
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    print("\n" + "=" * 65)
    print("  ⚠️ 시뮬레이션/분석용. 자동 매매 절대 금지.")
    print("=" * 65)

    return 0


if __name__ == "__main__":
    sys.exit(main())
