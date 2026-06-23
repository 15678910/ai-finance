"""
실적 모니터 — 분기 EPS·매출 실제 vs 컨센서스(서프라이즈)
========================================================
인베스팅닷컴 '실적' 탭 스타일: 발표일별 실제 EPS·매출 vs 예측, 서프라이즈%, 어닝 비트/미스.
대상: 마이크론(MU)·SK하이닉스·삼성전자. 데이터: yfinance(무료) — 최근 일부 분기 지연 가능.

출력: docs/earnings.json
🚨 통계·정보용. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "earnings.json")

# (표시명, 티커, 통화기호, EPS단위표기, 매출 나눔, 매출단위)
STOCKS = [
    ("마이크론", "MU", "$", "", 1e9, "B"),
    ("SK하이닉스", "000660.KS", "₩", "원", 1e12, "조"),
    ("삼성전자", "005930.KS", "₩", "원", 1e12, "조"),
]


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    import yfinance as yf
    import pandas as pd
    import warnings
    warnings.filterwarnings("ignore")

    now = datetime.now(KST)
    out_stocks = []

    for nm, tk, cur, eps_unit, rev_div, rev_unit in STOCKS:
        t = yf.Ticker(tk)
        # 분기 매출 맵: 'YYYY-MM' → 매출(원화/달러 raw)
        revmap = {}
        try:
            fin = t.quarterly_income_stmt
            if fin is not None and "Total Revenue" in fin.index:
                for c, v in fin.loc["Total Revenue"].dropna().items():
                    if v and float(v) > 0:
                        revmap[str(c)[:7]] = float(v)
        except Exception:
            pass

        def match_rev(dstr):
            """발표일(dstr) 직전 ~120일 내 분기말 매출."""
            try:
                d = datetime.strptime(dstr, "%Y-%m-%d")
            except Exception:
                return None
            best, bestgap = None, 999
            for k, v in revmap.items():
                try:
                    qd = datetime.strptime(k + "-15", "%Y-%m-%d")
                except Exception:
                    continue
                gap = (d - qd).days
                if 0 <= gap <= 130 and gap < bestgap:
                    best, bestgap = v, gap
            return best

        quarters = []
        try:
            ed = t.get_earnings_dates(limit=16)
        except Exception:
            ed = None
        if ed is not None and len(ed):
            seen = set()
            for idx, row in ed.iterrows():
                d = str(idx)[:10]
                if d in seen:
                    continue
                seen.add(d)
                ea = row.get("Reported EPS")
                ee = row.get("EPS Estimate")
                sp = row.get("Surprise(%)")
                ea = None if (ea is None or pd.isna(ea)) else float(ea)
                ee = None if (ee is None or pd.isna(ee)) else float(ee)
                sp = None if (sp is None or pd.isna(sp)) else float(sp)
                if ea is None and ee is None:
                    continue
                rev = match_rev(d)
                quarters.append({
                    "date": d,
                    "eps_actual": round(ea, 2) if ea is not None else None,
                    "eps_est": round(ee, 2) if ee is not None else None,
                    "surprise_pct": round(sp, 1) if sp is not None else None,
                    "beat": (ea is not None and ee is not None and ea >= ee),
                    "rev_actual": round(rev / rev_div, 2) if rev else None,
                    "upcoming": (ea is None and ee is not None),
                })
        quarters.sort(key=lambda q: q["date"], reverse=True)
        quarters = quarters[:8]

        # 적중 통계(실제 발표된 분기 중 컨센 상회 비율)
        graded = [q for q in quarters if q["eps_actual"] is not None and q["eps_est"] is not None]
        beat_rate = round(sum(1 for q in graded if q["beat"]) / len(graded) * 100) if graded else None
        avg_surprise = round(sum(q["surprise_pct"] for q in graded if q["surprise_pct"] is not None)
                             / max(1, len([q for q in graded if q["surprise_pct"] is not None])), 1) if graded else None

        out_stocks.append({
            "name": nm, "ticker": tk, "currency": cur, "eps_unit": eps_unit,
            "rev_unit": rev_unit, "quarters": quarters,
            "beat_rate": beat_rate, "avg_surprise_pct": avg_surprise, "n_graded": len(graded),
        })
        print(f"  {nm}: {len(quarters)}개 분기 · 컨센 상회 {beat_rate}% · 평균 서프라이즈 {avg_surprise}%")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "stocks": out_stocks,
        "note": ("분기 EPS·매출 실제 vs 컨센서스(서프라이즈%). yfinance 무료 데이터 — "
                 "최근 분기·매출 컨센서스는 지연/결측 가능. 비트=실제 EPS ≥ 예측. 통계용·투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
