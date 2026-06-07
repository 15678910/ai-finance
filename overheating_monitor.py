"""
과열·차익실현 위험 게이지
==========================
"1.9 고정 트리거" 같은 미신이 아니라, 통계적 과열도로 차익실현/조정 위험을 측정.

각 지수(SOX/나스닥100/KOSPI/EWY)에 대해:
  - z50: 50일 이동평균 대비 표준편차(z-score) — 단기 신전 정도
  - mom_z: 60일 모멘텀의 z-score — 추세 과열
  - disp200: 200일선 이격도(%) — 중기 과열
  → 종합 heat(0~100) + 🟢안전/🟡과열/🟠위험/🔴극단 신호등
  역사적 위험선: +3σ (이번 붕괴 직전 고점대)

출력: docs/overheating.json
🚨 통계 추정. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "overheating.json")
OVERSEAS_FILE = os.path.join(BASE_DIR, "docs", "overseas_market.json")

INDICES = [
    {"key": "SOX",   "ticker": "^SOX",  "name": "필라델피아 반도체", "flag": "🇺🇸"},
    {"key": "NDX",   "ticker": "^NDX",  "name": "나스닥100",        "flag": "🇺🇸"},
    {"key": "KOSPI", "ticker": "^KS11", "name": "KOSPI",            "flag": "🇰🇷"},
    {"key": "EWY",   "ticker": "EWY",   "name": "한국 ETF (외국인)", "flag": "🇰🇷"},
]

DANGER_Z = 3.0  # 이번 붕괴 직전 고점대 (~+3σ)


def classify(heat: float) -> tuple:
    if heat >= 85:
        return "극단(차익실현 위험)", "red"
    if heat >= 72:
        return "위험", "orange"
    if heat >= 60:
        return "과열", "amber"
    if heat >= 45:
        return "정상", "cyan"
    if heat >= 32:
        return "중립", "green"
    return "침체", "muted"


def get_vix() -> dict:
    try:
        with open(OVERSEAS_FILE, encoding="utf-8") as f:
            om = json.load(f)
        for m in om.get("all_markets", []):
            if m.get("ticker") == "^VIX" or m.get("is_vix"):
                return {"current": m.get("current"), "change_pct": m.get("change_pct")}
    except Exception:
        pass
    return {"current": None, "change_pct": None}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    import numpy as np
    import pandas as pd
    import yfinance as yf
    import warnings
    warnings.filterwarnings("ignore")

    print("=" * 55)
    print("  과열·차익실현 위험 게이지")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    tickers = [x["ticker"] for x in INDICES]
    raw = yf.download(tickers, period="2y", interval="1d", progress=False, auto_adjust=True)
    px = raw["Close"].dropna()
    asof = px.index[-1].strftime("%Y-%m-%d")

    results = []
    for idx in INDICES:
        s = px[idx["ticker"]].dropna()
        if len(s) < 210:
            print(f"  [WARN] {idx['key']} 데이터 부족")
            continue
        ma50 = s.rolling(50).mean()
        sd50 = s.rolling(50).std()
        z = (s - ma50) / sd50
        z50 = float(z.iloc[-1])
        ma200 = s.rolling(200).mean()
        disp200 = float((s.iloc[-1] / ma200.iloc[-1] - 1) * 100)
        ret60 = np.log(s).diff(60)
        mom_z_series = (ret60 - ret60.rolling(250).mean()) / ret60.rolling(250).std()
        mom_z = float(mom_z_series.iloc[-1]) if not np.isnan(mom_z_series.iloc[-1]) else 0.0
        # 최근 60일 고점 z
        peak_z = float(z.iloc[-60:].max())
        peak_date = z.iloc[-60:].idxmax().strftime("%Y-%m-%d")

        # 종합 heat (0~100): z50 주축 + 모멘텀 + 이격도
        heat = 50 + 13 * z50 + 5 * mom_z + 0.25 * min(max(disp200, 0), 80)
        heat = round(max(0, min(100, heat)), 1)
        signal, color = classify(heat)

        entry = {
            "key": idx["key"], "name": idx["name"], "flag": idx["flag"],
            "price": round(float(s.iloc[-1]), 2),
            "z50": round(z50, 2), "mom_z": round(mom_z, 2),
            "disp200": round(disp200, 1),
            "heat": heat, "signal": signal, "color": color,
            "peak_z": round(peak_z, 2), "peak_date": peak_date,
            "danger_z": DANGER_Z,
            "z_to_danger": round(DANGER_Z - z50, 2),
        }
        results.append(entry)
        print(f"  {idx['key']:6} heat={heat:5} {signal:14} | z50=+{z50:.2f}σ mom_z={mom_z:+.2f} 200이격={disp200:+.1f}% (고점 +{peak_z:.2f}σ {peak_date})")

    vix = get_vix()
    avg_heat = round(sum(r["heat"] for r in results) / len(results), 1) if results else None

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": asof,
        "vix": vix,
        "avg_heat": avg_heat,
        "danger_z": DANGER_Z,
        "indices": results,
        "method": "50일선 z-score(주축) + 60일 모멘텀 z + 200일 이격도 → heat 0~100",
        "myth_note": "'1.9 고정 트리거' 설은 데이터로 반증됨. 실제 고점은 ~+3σ, 트리거는 변동성·극단 이격(고정 가격 아님).",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료 (avg_heat={avg_heat})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
