"""
미국 신용 스프레드 + 거시 금리 모니터 (FRED 기반)
====================================================

목적:
  세인트루이스 연준 FRED에서 매일 신용 스프레드 + 국채 금리 데이터를 수집하여
  경기 침체 / 신용 시장 스트레스 / 위험 회피 강도를 추적.

추적 지표:
  · HY (High Yield) OAS: 하이일드 회사채 vs 국채 — 경기침체 선행지표
  · AAA OAS: 최고등급 회사채 vs 국채 — 안전자산 시장 스트레스
  · BBB OAS: 투자등급 하한 vs 국채 — 신용 분화 관찰
  · BBB-AAA: 투자등급 내부 신용 분화 — 위험회피 강도
  · DGS10 / DGS2: 10년·2년 국채 수익률
  · T10Y2Y: 10년-2년 스프레드 (역전 = 침체 6-18개월 선행)

데이터 소스:
  · https://fred.stlouisfed.org/graph/fredgraph.csv?id={SERIES_ID}
  · 무료 + API 키 불필요
  · 일별 갱신 (미국 영업일 기준)

추적 시그널:
  ⭐⭐⭐⭐⭐ HY OAS > 800bp = 경기침체 임박 신호
  ⭐⭐⭐⭐⭐ T10Y2Y 역전 (음수) = 침체 선행 신호
  ⭐⭐⭐⭐  HY OAS 30일 +100bp 이상 급등 = 위험 회피 시작
  ⭐⭐⭐⭐  BBB-AAA 스프레드 100bp+ = 신용 분화 (스트레스)
  ⭐⭐⭐   HY OAS 1년 90 백분위수+ = 역사적 고점 근접
  ⭐⭐⭐   T10Y2Y 5일 변화 -20bp+ = 곡선 평탄화 가속

🚨 정보 모니터링. 투자 결정 단독 사용 금지.
"""

import os
import io
import sys
import csv
import json
import time
import urllib.request
import urllib.error
from datetime import datetime, timezone, timedelta
from statistics import median

from core import send_message, load_state, save_state

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "credit_spread.json")
STATE_NAME = "credit_spread"
KST = timezone(timedelta(hours=9))

FRED_BASE = "https://fred.stlouisfed.org/graph/fredgraph.csv"
USER_AGENT = "Mozilla/5.0 (compatible; ai-finance-credit-monitor/1.0)"
TIMEOUT_SEC = 30

# ====================================================================
# 추적 시리즈 정의
# ====================================================================
SERIES = [
    # 신용 스프레드 (OAS) — 단위: %p (FRED는 %로 반환)
    {
        "id": "BAMLH0A0HYM2",
        "name": "하이일드 OAS",
        "category": "credit_spread",
        "subcategory": "high_yield",
        "unit": "%p",
        "description": "ICE BofA US High Yield OAS — 하이일드 채권 vs 동기간 국채 스프레드",
        "thresholds": {"caution": 5.0, "warning": 7.0, "recession": 8.0},
    },
    {
        "id": "BAMLC0A1CAAA",
        "name": "AAA OAS",
        "category": "credit_spread",
        "subcategory": "investment_grade",
        "unit": "%p",
        "description": "ICE BofA AAA Corporate OAS — 최우량 회사채 스프레드",
        "thresholds": {"caution": 0.7, "warning": 1.0, "recession": 1.5},
    },
    {
        "id": "BAMLC0A4CBBB",
        "name": "BBB OAS",
        "category": "credit_spread",
        "subcategory": "investment_grade",
        "unit": "%p",
        "description": "ICE BofA BBB Corporate OAS — 투자등급 하한 스프레드",
        "thresholds": {"caution": 2.0, "warning": 2.5, "recession": 3.5},
    },
    {
        "id": "BAMLC0A0CM",
        "name": "투자등급 OAS",
        "category": "credit_spread",
        "subcategory": "investment_grade",
        "unit": "%p",
        "description": "ICE BofA US Corporate OAS — 투자등급 전체 평균",
        "thresholds": {"caution": 1.5, "warning": 2.0, "recession": 2.8},
    },
    # 국채 금리 + 곡선 — 단위: %
    {
        "id": "DGS10",
        "name": "10Y 국채 수익률",
        "category": "yield",
        "unit": "%",
        "description": "Market Yield on US Treasury Securities at 10-Year Constant Maturity",
    },
    {
        "id": "DGS2",
        "name": "2Y 국채 수익률",
        "category": "yield",
        "unit": "%",
        "description": "Market Yield on US Treasury Securities at 2-Year Constant Maturity",
    },
    {
        "id": "T10Y2Y",
        "name": "10Y-2Y 스프레드",
        "category": "yield_curve",
        "unit": "%p",
        "description": "10-Year minus 2-Year Treasury (음수 = 장단기 역전 = 침체 선행)",
        "thresholds": {"recession": -0.0001},  # 0 미만 = 역전
    },
]


# ====================================================================
# FRED CSV 다운로드
# ====================================================================
def fetch_fred_series(series_id: str, since: str = "2020-01-01") -> list:
    """FRED CSV → [{date, value}] (역사적 데이터)."""
    url = f"{FRED_BASE}?id={series_id}&cosd={since}"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(req, timeout=TIMEOUT_SEC) as resp:  # nosec — 공개 데이터
            content = resp.read().decode("utf-8")
    except Exception as e:
        print(f"    [실패] {series_id}: {e}")
        return []

    rows = []
    reader = csv.reader(io.StringIO(content))
    next(reader, None)  # header skip
    for line in reader:
        if len(line) < 2:
            continue
        date_str, value_str = line[0].strip(), line[1].strip()
        if value_str in (".", "", "N/A"):
            continue
        try:
            value = float(value_str)
            rows.append({"date": date_str, "value": value})
        except ValueError:
            continue
    return rows


# ====================================================================
# 분석: 변화율, 백분위수, 시그널
# ====================================================================
def analyze_series(series_def: dict, data: list) -> dict:
    """시계열 데이터 분석 → 시그널 + 통계."""
    if not data or len(data) < 30:
        return {"insufficient_data": True}

    # 최신순으로 정렬
    sorted_data = sorted(data, key=lambda r: r["date"], reverse=True)
    latest = sorted_data[0]
    values_latest = [r["value"] for r in sorted_data]

    # 변화량 계산
    def change_at(days_back: int) -> float | None:
        if len(sorted_data) <= days_back:
            return None
        return latest["value"] - sorted_data[days_back]["value"]

    def value_at(days_back: int) -> float | None:
        if len(sorted_data) <= days_back:
            return None
        return sorted_data[days_back]["value"]

    change_1d = change_at(1)
    change_5d = change_at(5)
    change_30d = change_at(30)
    change_1y = change_at(252)

    # 백분위수 (최근 1년, 5년)
    recent_1y = values_latest[:252] if len(values_latest) >= 252 else values_latest
    recent_5y = values_latest[:1260] if len(values_latest) >= 1260 else values_latest

    def percentile_of(values, target):
        if not values:
            return None
        sorted_vals = sorted(values)
        lower = sum(1 for v in sorted_vals if v < target)
        return round(lower / len(sorted_vals) * 100, 1)

    pct_1y = percentile_of(recent_1y, latest["value"])
    pct_5y = percentile_of(recent_5y, latest["value"])

    # 평균
    avg_1y = round(sum(recent_1y) / len(recent_1y), 3) if recent_1y else None
    median_1y = round(median(recent_1y), 3) if recent_1y else None
    max_1y = max(recent_1y) if recent_1y else None
    min_1y = min(recent_1y) if recent_1y else None

    # 시그널 감지
    signals = []
    thresholds = series_def.get("thresholds", {})
    current = latest["value"]

    # 1. 임계값 초과 (HY recession 등)
    if "recession" in thresholds:
        if series_def["id"] == "T10Y2Y":
            # 역전 신호 (음수)
            if current < 0:
                signals.append({
                    "type": "yield_curve_inverted",
                    "strength": 5,
                    "emoji": "🚨",
                    "label": "장단기 금리 역전",
                    "description": f"T10Y2Y = {current:.2f}%p (음수 = 침체 6~18개월 선행 시그널)",
                })
        elif current >= thresholds["recession"]:
            signals.append({
                "type": "recession_level",
                "strength": 5,
                "emoji": "🚨",
                "label": f"{series_def['name']} 침체 수준",
                "description": f"현재 {current:.2f}{series_def['unit']} ≥ 침체 임계 {thresholds['recession']:.2f}",
            })
        elif "warning" in thresholds and current >= thresholds["warning"]:
            signals.append({
                "type": "warning_level",
                "strength": 4,
                "emoji": "⚠️",
                "label": f"{series_def['name']} 경고 수준",
                "description": f"현재 {current:.2f}{series_def['unit']} ≥ 경고 {thresholds['warning']:.2f}",
            })
        elif "caution" in thresholds and current >= thresholds["caution"]:
            signals.append({
                "type": "caution_level",
                "strength": 3,
                "emoji": "🟡",
                "label": f"{series_def['name']} 주의 수준",
                "description": f"현재 {current:.2f}{series_def['unit']} ≥ 주의 {thresholds['caution']:.2f}",
            })

    # 2. 급변
    if change_30d is not None and series_def["category"] == "credit_spread":
        if change_30d >= 1.0:
            signals.append({
                "type": "spread_widening_fast",
                "strength": 4,
                "emoji": "📈",
                "label": "30일 스프레드 급등",
                "description": f"30일간 +{change_30d:.2f}%p (위험 회피 가속)",
            })
        elif change_30d <= -1.0:
            signals.append({
                "type": "spread_tightening_fast",
                "strength": 3,
                "emoji": "📉",
                "label": "30일 스프레드 급락",
                "description": f"30일간 {change_30d:.2f}%p (위험 선호 회복)",
            })

    # 3. 역사적 백분위 극단
    if pct_1y is not None:
        if pct_1y >= 95:
            signals.append({
                "type": "near_1y_high",
                "strength": 3,
                "emoji": "🔝",
                "label": "1년 최고치 근접",
                "description": f"1년 백분위 {pct_1y:.0f}% (역사적 고점 권역)",
            })
        elif pct_1y <= 5:
            signals.append({
                "type": "near_1y_low",
                "strength": 3,
                "emoji": "🔻",
                "label": "1년 최저치 근접",
                "description": f"1년 백분위 {pct_1y:.0f}% (역사적 저점 권역)",
            })

    return {
        "latest_date": latest["date"],
        "latest_value": round(current, 3),
        "unit": series_def["unit"],
        "change_1d": round(change_1d, 3) if change_1d is not None else None,
        "change_5d": round(change_5d, 3) if change_5d is not None else None,
        "change_30d": round(change_30d, 3) if change_30d is not None else None,
        "change_1y": round(change_1y, 3) if change_1y is not None else None,
        "percentile_1y": pct_1y,
        "percentile_5y": pct_5y,
        "avg_1y": avg_1y,
        "median_1y": median_1y,
        "max_1y": round(max_1y, 3) if max_1y is not None else None,
        "min_1y": round(min_1y, 3) if min_1y is not None else None,
        "signals": signals,
        "max_strength": max((s["strength"] for s in signals), default=0),
    }


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 72)
    print("  미국 신용 스프레드 + 거시 금리 모니터 (FRED)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  대상: {len(SERIES)}개 시리즈")
    print("=" * 72)

    state = load_state(STATE_NAME, default={"last_alerted_signals": {}})
    last_alerted = state.get("last_alerted_signals", {})

    # 시리즈별 수집 + 분석
    print("\n[FRED 데이터 수집]")
    results = []
    for s in SERIES:
        print(f"  {s['name']} ({s['id']})...", end=" ")
        data = fetch_fred_series(s["id"])
        if not data:
            print("실패")
            continue
        analysis = analyze_series(s, data)
        results.append({
            "id": s["id"],
            "name": s["name"],
            "category": s["category"],
            "subcategory": s.get("subcategory"),
            "unit": s["unit"],
            "description": s["description"],
            "fred_url": f"https://fred.stlouisfed.org/series/{s['id']}",
            "recent_30d": data[-30:],  # 차트용
            **analysis,
        })
        latest_v = analysis.get("latest_value", "—")
        ch_30d = analysis.get("change_30d", "—")
        print(f"{latest_v}{s['unit']} (30d {ch_30d:+}{s['unit'] if ch_30d != '—' else ''})")
        time.sleep(0.3)  # rate limit

    # 강력 시그널 종합
    strong_signals = []
    for r in results:
        for s in r.get("signals", []):
            if s["strength"] >= 4:
                strong_signals.append({
                    "series_id": r["id"],
                    "series_name": r["name"],
                    "latest_value": r.get("latest_value"),
                    "unit": r["unit"],
                    **s,
                })
    strong_signals.sort(key=lambda x: x["strength"], reverse=True)

    # 신규 알림 (이전 알림과 다른 시그널만)
    new_alerts = []
    current_signal_keys = {}
    for sig in strong_signals:
        sig_key = f"{sig['series_id']}:{sig['type']}"
        current_signal_keys[sig_key] = sig["latest_value"]
        # 이전에 알림 안 했거나, 값이 5% 이상 변했으면 재알림
        prev_value = last_alerted.get(sig_key)
        if prev_value is None or abs(sig["latest_value"] - prev_value) / max(abs(prev_value), 0.01) > 0.05:
            new_alerts.append(sig)

    if new_alerts:
        lines = ["💳 미국 신용 시장 · 경기 시그널", "=" * 30, ""]
        for sig in new_alerts[:8]:
            lines.append(f"{sig['emoji']} {sig['series_name']}: {sig['label']}")
            lines.append(f"   {sig['description']}")
            lines.append("")
        lines.append("🚨 정보 모니터링. 투자 결정 단독 사용 금지.")
        lines.append("대시보드: https://15678910.github.io/ai-finance/")
        try:
            send_message("\n".join(lines))
            print(f"\n  ✅ 텔레그램 발송: 신규 시그널 {len(new_alerts)}건")
        except Exception as e:
            print(f"\n  ❌ 텔레그램 발송 실패: {e}")

    state["last_alerted_signals"] = current_signal_keys
    save_state(STATE_NAME, state)

    # 종합 거시 등급 (HY OAS 기준)
    hy = next((r for r in results if r["id"] == "BAMLH0A0HYM2"), None)
    t10y2y = next((r for r in results if r["id"] == "T10Y2Y"), None)

    macro_regime = "정상"
    macro_regime_color = "green"
    macro_regime_emoji = "🟢"
    if hy and t10y2y:
        hy_v = hy.get("latest_value", 0)
        curve_v = t10y2y.get("latest_value", 0)
        if (hy_v and hy_v >= 8) or (curve_v is not None and curve_v < -0.5):
            macro_regime = "침체 임박 신호"
            macro_regime_color = "red"
            macro_regime_emoji = "🚨"
        elif (hy_v and hy_v >= 6) or (curve_v is not None and curve_v < 0):
            macro_regime = "경계 모드 (장단기 역전 / HY 경고)"
            macro_regime_color = "amber"
            macro_regime_emoji = "⚠️"
        elif (hy_v and hy_v >= 5):
            macro_regime = "주의 (HY 상승)"
            macro_regime_color = "yellow"
            macro_regime_emoji = "🟡"

    # 결과 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "methodology": (
            "FRED CSV 무인증 다운로드 → 신용 스프레드 + 국채 금리 일별 추적. "
            "백분위·변화율·임계값 기반 시그널 자동 감지."
        ),
        "macro_regime": macro_regime,
        "macro_regime_color": macro_regime_color,
        "macro_regime_emoji": macro_regime_emoji,
        "series_count": len(results),
        "strong_signals": strong_signals,
        "strong_signals_count": len(strong_signals),
        "new_alerts_this_run": len(new_alerts),
        "results": results,
        "interpretation_guide": {
            "HY_OAS": "5%p+ = 주의 / 7%p+ = 경고 / 8%p+ = 침체 임박",
            "BBB_AAA_spread": "확대 = 투자등급 내부 신용 분화 = 위험회피",
            "T10Y2Y": "음수 = 장단기 역전 = 6~18개월 후 침체 가능성",
        },
        "data_source": "FRED (St. Louis Fed) — fred.stlouisfed.org",
        "warning": "🚨 거시 정보 모니터링. 투자 결정 단독 사용 금지.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    print(f"\n[거시 체제] {macro_regime_emoji} {macro_regime}")
    print(f"\n[강력 시그널 Top 5]")
    for sig in strong_signals[:5]:
        print(f"  {sig['emoji']} {sig['series_name']}: {sig['label']}")

    print("\n" + "=" * 72)
    print(f"  완료: {len(results)}/{len(SERIES)} 시리즈 · 시그널 {len(strong_signals)}건 · 신규 알림 {len(new_alerts)}건")
    print("=" * 72)


if __name__ == "__main__":
    main()
