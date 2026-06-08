"""
외국인 보유 + 개인 추정 평단가
==============================
- 외국인 보유율·보유주식수: 정확 (KRX/네이버 공시)
- 개인 추정 평단가: 순매수 가중평균 (60일 우선, 없으면 10일) — 추정치(공시 없음)
- 평가손익 = 현재가 / 추정평단가 - 1

데이터: docs/data.json (stock_investor_details 60일 + investor_flow 46종목)
출력: docs/holdings_estimate.json
🚨 평단가는 추정치(실제 공시 없음). 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "docs", "data.json")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "holdings_estimate.json")

# 주요 종목 (반도체·테크·유통)
TARGETS = ["005930", "000660", "066570", "005380", "042700",
           "035420", "035720", "139480", "004170", "023530"]


def _weighted_avg(rows, net_key, close_key="close"):
    """순매수일(net>0)만 가중평균단가 + 순매수일수 + 누적순매수."""
    num = den = 0.0
    buy_days = 0
    cum = 0.0
    for r in rows:
        net = r.get(net_key)
        c = r.get(close_key)
        if not isinstance(net, (int, float)) or not isinstance(c, (int, float)):
            continue
        cum += net
        if net > 0:
            num += net * c
            den += net
            buy_days += 1
    return (round(num / den) if den else None), buy_days, round(cum)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    print("=" * 55)
    print("  외국인 보유 + 개인 추정 평단가")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    with open(DATA_FILE, encoding="utf-8") as f:
        data = json.load(f)

    # 현재가: yfinance 실시간 (60일/플로 데이터는 stale일 수 있음)
    def live_price(tk):
        try:
            import yfinance as yf
            for suf in (".KS", ".KQ"):
                p = getattr(yf.Ticker(tk + suf).fast_info, "last_price", None)
                if p:
                    return round(float(p))
        except Exception:
            pass
        return None

    # 60일 소스 (stock_investor_details) — individual_net_shares_est 보유
    sid = {s["ticker"]: s for s in (data.get("stock_investor_details", {}) or {}).get("stocks", [])}
    # 10일 소스 (investor_flow) — 전 종목, 개인 ≈ -(외인+기관)
    flow = {r["ticker"]: r for r in (data.get("investor_flow", {}) or {}).get("results", [])}

    results = []
    for tk in TARGETS:
        s60 = sid.get(tk)
        sf = flow.get(tk)
        name = (s60 or sf or {}).get("name", tk)
        if not (s60 or sf):
            continue

        WIN = 20  # 최근 20거래일(~1개월) 순매수 기준 — 현재 평가손익에 적합
        # 60일 데이터 우선 (개인 순매수 직접 추정값 보유)
        if s60 and s60.get("rows"):
            asc = s60["rows"] if s60["rows"][0].get("date", "") < s60["rows"][-1].get("date", "") else list(reversed(s60["rows"]))
            rows = asc[-WIN:]  # 최근 WIN일
            avg, buy_days, cum = _weighted_avg(rows, "individual_net_shares_est")
            src = f"최근{len(rows)}일"
            last = rows[-1]
        elif sf and sf.get("rows"):
            # investor_flow rows(최신순): 개인 = -(외인+기관)
            rows = []
            for r in sf["rows"][:WIN]:
                fn = r.get("foreign_net_shares") or 0
                inn = r.get("institutional_net_shares") or 0
                rows.append({**r, "ind": -(fn + inn)})
            avg, buy_days, cum = _weighted_avg(rows, "ind")
            src = f"최근{len(rows)}일"
            last = sf["rows"][0]  # flow rows는 최신순
        else:
            continue

        # 현재가: yfinance 실시간 우선, 실패 시 가장 신선한 embedded close
        cur = live_price(tk) or (sf or {}).get("latest_close") or last.get("close")
        fpct = last.get("foreign_holding_pct")
        fsh = last.get("foreign_holding_shares")
        if fpct is None and sf:
            fpct = sf.get("foreign_holding_pct_now")
        pnl = round((cur / avg - 1) * 100, 1) if (avg and cur) else None
        f5d = sf.get("foreign_holding_pct_5d_change") if sf else None
        streak = (sf.get("foreign_streak_days"), sf.get("foreign_streak_direction")) if sf else (None, None)

        entry = {
            "ticker": tk, "name": name,
            "current": cur,
            "foreign_pct": fpct,
            "foreign_shares": fsh,
            "foreign_pct_5d_change": f5d,
            "foreign_streak_days": streak[0],
            "foreign_streak_dir": streak[1],
            "indiv_avg_est": avg,
            "indiv_pnl_pct": pnl,
            "indiv_buy_days": buy_days,
            "indiv_cum_net": cum,
            "window": src,
        }
        results.append(entry)
        print(f"  {name}({tk}): 외인 {fpct}% | 개인추정평단 {avg} | 손익 {pnl}% | {src}")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "stocks": results,
        "note": "외국인 보유율·보유량=공시(정확). 개인 평단가=순매수 가중평균 추정(공시 없음). 평가손익=현재가/추정평단-1.",
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} ({len(results)}종목)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
