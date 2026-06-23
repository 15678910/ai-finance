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

        # ── #1·#2 보완: 예정(다음) 분기 컨센서스 (선행 EPS·매출 추정) ──
        next_earn = None
        try:
            ee = t.earnings_estimate          # 0q=이번(예정) 분기
            rev = t.revenue_estimate
            eps_e = float(ee.loc["0q", "avg"]) if (ee is not None and "0q" in ee.index and "avg" in ee.columns) else None
            rev_e = float(rev.loc["0q", "avg"]) if (rev is not None and "0q" in rev.index and "avg" in rev.columns) else None
            nd_raw = None
            try:
                cal = t.calendar or {}
                ed_cal = cal.get("Earnings Date")
                if ed_cal:
                    nd_raw = str(ed_cal[0])[:10] if isinstance(ed_cal, (list, tuple)) else str(ed_cal)[:10]
            except Exception:
                pass
            # 발표일 검증: 미래면 확정, 과거/결측이면 분기(~91일) 주기로 추정 — '과거를 예정'으로 표시 방지
            today_d = now.date()
            nd_final, nd_status = None, "tbd"
            if nd_raw:
                try:
                    if datetime.strptime(nd_raw, "%Y-%m-%d").date() >= today_d:
                        nd_final, nd_status = nd_raw, "confirmed"
                except Exception:
                    pass
            act_dates = [str(q["date"]) for q in quarters if len(str(q.get("date", ""))) == 10]
            if nd_final is None and act_dates:
                try:
                    est = datetime.strptime(max(act_dates), "%Y-%m-%d").date()
                    while est < today_d:
                        est = est + timedelta(days=91)
                    nd_final, nd_status = est.strftime("%Y-%m-%d"), "estimated"
                except Exception:
                    pass
            # 실제 분기 지연(최신 실제 발표가 150일+ 과거 = 무료 피드 반영 대기)
            stale_q = False
            if act_dates:
                try:
                    stale_q = (today_d - datetime.strptime(max(act_dates), "%Y-%m-%d").date()).days > 150
                except Exception:
                    pass
            if eps_e is not None or rev_e is not None:
                next_earn = {"date": nd_final, "date_status": nd_status, "stale_actuals": stale_q,
                             "eps_est": round(eps_e, 2) if eps_e is not None else None,
                             "rev_est": round(rev_e / rev_div, 2) if rev_e else None}
                if not any(q.get("upcoming") for q in quarters):
                    quarters.insert(0, {"date": nd_final, "date_status": nd_status, "eps_actual": None,
                                        "eps_est": next_earn["eps_est"], "surprise_pct": None,
                                        "beat": False, "rev_actual": None, "rev_est": next_earn["rev_est"],
                                        "upcoming": True})
        except Exception:
            pass

        # ── 배당 이력 ──
        dividends = []
        ttm_div = 0.0
        try:
            dv = t.dividends
            if dv is not None and len(dv):
                from datetime import datetime as _dt
                recent = list(dv.items())[-8:]
                dividends = [{"date": str(d)[:10], "amount": round(float(v), 2)} for d, v in recent]
                # TTM 합계(최근 4건 근사)
                ttm_div = round(sum(float(v) for _, v in list(dv.items())[-4:]), 2)
        except Exception:
            pass
        # 배당수익률(TTM배당 ÷ 현재가)
        div_yield = None
        try:
            fi = t.fast_info
            px = float(getattr(fi, "last_price", 0) or 0)
            if px > 0 and ttm_div > 0:
                div_yield = round(ttm_div / px * 100, 2)
        except Exception:
            pass

        # ── 옵션 체인 (미국 종목만 — 한국 주식은 무료 피드 미제공) ──
        options = None
        try:
            exps = t.options
            if exps:
                exp = exps[0]
                oc = t.option_chain(exp)
                px = None
                try:
                    px = float(getattr(t.fast_info, "last_price", 0) or 0)
                except Exception:
                    pass

                def near(df, is_call):
                    df = df.sort_values("strike")
                    if px:
                        df["d"] = (df["strike"] - px).abs()
                        atm_i = df["d"].idxmin()
                        pos = df.index.get_loc(atm_i)
                        df = df.iloc[max(0, pos - 4):pos + 5]
                    else:
                        df = df.iloc[len(df) // 2 - 4:len(df) // 2 + 5]
                    rows = []
                    for _, r in df.iterrows():
                        rows.append({"strike": round(float(r["strike"]), 1),
                                     "last": round(float(r.get("lastPrice", 0) or 0), 2),
                                     "bid": round(float(r.get("bid", 0) or 0), 2),
                                     "ask": round(float(r.get("ask", 0) or 0), 2),
                                     "vol": int(r.get("volume", 0) or 0),
                                     "oi": int(r.get("openInterest", 0) or 0),
                                     "iv": round(float(r.get("impliedVolatility", 0) or 0) * 100, 1)})
                    return rows
                options = {"expiry": exp, "spot": round(px, 2) if px else None,
                           "calls": near(oc.calls, True), "puts": near(oc.puts, False),
                           "n_expiries": len(exps)}
        except Exception as e:
            print(f"  [WARN] {nm} 옵션 실패: {e}")

        # 적중 통계(실제 발표된 분기 중 컨센 상회 비율)
        graded = [q for q in quarters if q["eps_actual"] is not None and q["eps_est"] is not None]
        beat_rate = round(sum(1 for q in graded if q["beat"]) / len(graded) * 100) if graded else None
        avg_surprise = round(sum(q["surprise_pct"] for q in graded if q["surprise_pct"] is not None)
                             / max(1, len([q for q in graded if q["surprise_pct"] is not None])), 1) if graded else None

        out_stocks.append({
            "name": nm, "ticker": tk, "currency": cur, "eps_unit": eps_unit,
            "rev_unit": rev_unit, "quarters": quarters,
            "beat_rate": beat_rate, "avg_surprise_pct": avg_surprise, "n_graded": len(graded),
            "next_earnings": next_earn,
            "dividends": dividends, "ttm_div": ttm_div, "div_yield_pct": div_yield,
            "options": options,
        })
        print(f"  {nm}: 분기 {len(quarters)} · 비트 {beat_rate}% · 배당 {len(dividends)}건(수익률 {div_yield}%) · 옵션 {'있음' if options else '없음'}")

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
