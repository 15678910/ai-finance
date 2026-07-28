"""
추세추종 신호 보드 — Supertrend·EMA200·ADX + 5년 백테스트 + 몬테카를로 검증
==========================================================================
알고트레이딩(Jesse) 방법론을 한국 시장 '일봉'에 적용한 정보 패널용 수집기.
  · 대추세: 종가 > EMA200 (앵커)
  · 신호: Supertrend(10, 3) 방향 (트레일링 스탑 라인 겸용)
  · 확인: ADX(14) ≥ 20 (추세 강도)
  · 풀백: 상승추세 중 종가 < EMA20 (눌림 구간)
백테스트(롱온리): Supertrend 상향 전환 + 종가>EMA200 + ADX≥20 진입,
  Supertrend 하향 전환 시 청산(트레일링). 비용 왕복 0.28%(수수료+거래세) 반영.
몬테카를로(랜덤진입 검정): 같은 횟수·같은 보유기간의 무작위 진입 1,000회
  null 분포에서 실제 성과의 백분위 → '전략이 운보다 나은가' 판별.

🚨 매매 실행 없음 — 신호 정보 제공용. 과거 성과는 미래를 보장하지 않음. 투자자문 아님.
출력: docs/trend_signal.json
"""

import json
import os
import random
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "trend_signal.json")

TICKERS = [
    ("^KS11", "KOSPI", "지수(참고)"),
    ("000660.KS", "SK하이닉스", "보유"),
    ("005930.KS", "삼성전자", "관심"),
]
ST_PERIOD, ST_MULT, ADX_MIN = 10, 3.0, 20
COST = 0.0028                         # 왕복 비용(수수료 0.05%×2 + 거래세 0.18%)


def ema(vals, n):
    k, out, e = 2 / (n + 1), [], None
    for v in vals:
        e = v if e is None else v * k + e * (1 - k)
        out.append(e)
    return out


def atr_wilder(h, l, c, n):
    trs = [h[0] - l[0]]
    for i in range(1, len(c)):
        trs.append(max(h[i] - l[i], abs(h[i] - c[i - 1]), abs(l[i] - c[i - 1])))
    out, a = [], None
    for i, tr in enumerate(trs):
        if i < n:
            a = tr if a is None else (a * i + tr) / (i + 1)     # 초기 단순평균
        else:
            a = (a * (n - 1) + tr) / n                          # Wilder smoothing
        out.append(a)
    return out


def supertrend(h, l, c, n=10, mult=3.0):
    """(direction[+1/-1] 리스트, 스탑라인 리스트)."""
    atr = atr_wilder(h, l, c, n)
    up = [(h[i] + l[i]) / 2 + mult * atr[i] for i in range(len(c))]
    dn = [(h[i] + l[i]) / 2 - mult * atr[i] for i in range(len(c))]
    f_up, f_dn = up[:], dn[:]
    for i in range(1, len(c)):
        f_up[i] = min(up[i], f_up[i - 1]) if c[i - 1] <= f_up[i - 1] else up[i]
        f_dn[i] = max(dn[i], f_dn[i - 1]) if c[i - 1] >= f_dn[i - 1] else dn[i]
    d, line = [1] * len(c), [f_dn[0]] * len(c)
    for i in range(1, len(c)):
        if d[i - 1] == 1:
            d[i] = -1 if c[i] < f_dn[i] else 1
        else:
            d[i] = 1 if c[i] > f_up[i] else -1
        line[i] = f_dn[i] if d[i] == 1 else f_up[i]
    return d, line


def adx_wilder(h, l, c, n=14):
    pdm, ndm = [0.0], [0.0]
    for i in range(1, len(c)):
        um, dm = h[i] - h[i - 1], l[i - 1] - l[i]
        pdm.append(um if (um > dm and um > 0) else 0.0)
        ndm.append(dm if (dm > um and dm > 0) else 0.0)
    atr = atr_wilder(h, l, c, n)

    def wsm(xs):
        out, s = [], None
        for i, x in enumerate(xs):
            s = x if s is None else s - s / n + x
            out.append(s)
        return out
    spdm, sndm = wsm(pdm), wsm(ndm)
    dx = []
    for i in range(len(c)):
        pdi = (spdm[i] / (atr[i] * n) * 100) if atr[i] else 0
        ndi = (sndm[i] / (atr[i] * n) * 100) if atr[i] else 0
        dx.append(abs(pdi - ndi) / (pdi + ndi) * 100 if (pdi + ndi) else 0)
    out, a = [], None
    for i, x in enumerate(dx):
        a = x if a is None else (a * (n - 1) + x) / n
        out.append(a)
    return out


def backtest(c, dates, st_dir, e200, adx):
    """롱온리: ST 롱 + 종가>EMA200 + ADX≥20 '조건 충족 시' 진입 / ST 하향 전환 청산(트레일링).
    자산곡선(전략 vs 바이앤홀드 벤치마크)·수중기간·기대값 포함 — Jesse 스타일 리포트용.
    ※ '전환 당일'로 제한하면 ST가 이미 롱인 채 대추세 조건이 갖춰지는 상승장 대부분을 놓침."""
    n = len(c)
    if n <= 210:
        return None
    trades, pos, entry_eq = [], None, 1.0
    eq = 1.0
    curve_eq, curve_bh, curve_d = [], [], []
    bh0 = c[201]
    for i in range(201, n):
        if pos is None:
            if st_dir[i] == 1 and c[i] > e200[i] and adx[i] >= ADX_MIN:
                pos, entry_eq = i, eq
        else:
            eq *= c[i] / c[i - 1]
            if st_dir[i] == -1:
                eq *= (1 - COST)
                trades.append({"ret": eq / entry_eq - 1, "days": i - pos})
                pos = None
        curve_eq.append(eq)
        curve_bh.append(c[i] / bh0)
        curve_d.append(dates[i])
    open_trade = False
    if pos is not None:
        trades.append({"ret": eq / entry_eq - 1, "days": n - 1 - pos, "open": True})
        open_trade = True
    if not trades:
        return None

    rets = [t["ret"] for t in trades]
    wins = [r for r in rets if r > 0]
    losses = [r for r in rets if r <= 0]
    years = (n - 201) / 252
    # 샤프: 자산곡선 일간수익 전체(현금 보유일 0% 포함)
    dr = [curve_eq[i] / curve_eq[i - 1] - 1 for i in range(1, len(curve_eq))]
    sharpe = None
    if len(dr) > 20:
        mu = sum(dr) / len(dr)
        sd = (sum((x - mu) ** 2 for x in dr) / (len(dr) - 1)) ** 0.5
        sharpe = round(mu / sd * (252 ** 0.5), 2) if sd else None
    # MDD·최대 수중기간(자산곡선 기준)
    peak, mdd, uw, max_uw = curve_eq[0], 0.0, 0, 0
    for v in curve_eq:
        if v >= peak:
            peak, uw = v, 0
        else:
            uw += 1
            max_uw = max(max_uw, uw)
            mdd = min(mdd, v / peak - 1)
    avg_win = sum(wins) / len(wins) * 100 if wins else None
    avg_loss = sum(losses) / len(losses) * 100 if losses else None
    wl = round(avg_win / abs(avg_loss), 2) if (avg_win is not None and avg_loss) else None
    # 차트용 다운샘플(~320점)
    step = max(1, len(curve_eq) // 320)
    samp = list(range(0, len(curve_eq), step))
    if samp[-1] != len(curve_eq) - 1:
        samp.append(len(curve_eq) - 1)
    return {
        "n_trades": len(trades), "n_open": 1 if open_trade else 0,
        "win_rate": round(len(wins) / len(trades) * 100),
        "avg_win_pct": round(avg_win, 2) if avg_win is not None else None,
        "avg_loss_pct": round(avg_loss, 2) if avg_loss is not None else None,
        "wl_ratio": wl,
        "expectancy_pct": round(sum(rets) / len(rets) * 100, 2),
        "total_return_pct": round((eq - 1) * 100, 1),
        "bench_return_pct": round((curve_bh[-1] - 1) * 100, 1),
        "cagr_pct": round(((eq ** (1 / years)) - 1) * 100, 1) if years > 0.5 else None,
        "sharpe": sharpe, "mdd_pct": round(mdd * 100, 1), "max_underwater_days": max_uw,
        "trades_per_year": round(len(trades) / years, 1) if years > 0.5 else None,
        "avg_hold_days": round(sum(t["days"] for t in trades) / len(trades)),
        "curve": {"d0": curve_d[0], "d1": curve_d[-1],
                  "eq": [round(curve_eq[i], 4) for i in samp],
                  "bh": [round(curve_bh[i], 4) for i in samp]},
        "durations": [t["days"] for t in trades],
    }


def monte_carlo(c, durations, actual_total_pct, sims=1000, seed=7):
    """랜덤진입 검정: 같은 횟수·보유기간의 무작위 진입 null 분포 → 실제 성과 백분위."""
    rng = random.Random(seed)
    n = len(c)
    nulls = []
    for _ in range(sims):
        tot = 1.0
        for d in durations:
            s = rng.randrange(201, n - d) if n - d > 201 else 201
            tot *= (1 + (c[s + d] / c[s] - 1 - COST))
        nulls.append((tot - 1) * 100)
    nulls.sort()
    below = sum(1 for x in nulls if x < actual_total_pct)
    pctile = round(below / sims * 100, 1)
    med = nulls[sims // 2]
    if pctile >= 95:
        v, col = "운 아님 — 유의미한 엣지 ✅", "green"
    elif pctile >= 80:
        v, col = "무작위보다 우수 (경계선)", "yellow"
    else:
        v, col = "무작위와 구분 불가", "red"
    return {"pctile": pctile, "null_median_pct": round(med, 1), "sims": sims,
            "verdict": v, "verdict_color": col}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    try:
        import yfinance as yf
    except Exception as e:
        print(f"[ERROR] yfinance 없음: {e}")
        return 1

    now = datetime.now(KST)
    out_t, asof = [], None
    for sym, name, tag in TICKERS:
        try:
            df = yf.Ticker(sym).history(period="5y", interval="1d", auto_adjust=True)
        except Exception as e:
            print(f"  [WARN] {name} 다운로드 실패: {e}")
            continue
        if df is None or len(df) < 260:
            print(f"  [WARN] {name} 데이터 부족")
            continue
        h, l, c = df["High"].tolist(), df["Low"].tolist(), df["Close"].tolist()
        dts = [x.strftime("%Y-%m-%d") for x in df.index]
        e200, e20 = ema(c, 200), ema(c, 20)
        st_dir, st_line = supertrend(h, l, c, ST_PERIOD, ST_MULT)
        adx = adx_wilder(h, l, c, 14)
        asof = df.index[-1].strftime("%Y-%m-%d")

        up = c[-1] > e200[-1]
        st_long = st_dir[-1] == 1
        strong = adx[-1] >= ADX_MIN
        pullback = up and st_long and c[-1] < e20[-1]
        if up and st_long and strong:
            verdict, vcol = ("추세추종 조건 충족 (풀백 중)" if pullback else "추세추종 조건 충족"), "green"
        elif up and st_long:
            verdict, vcol = "상승추세 · 추세강도 약함(ADX<20)", "yellow"
        elif up:
            verdict, vcol = "대추세 상승 · Supertrend 하향(조정 중)", "yellow"
        else:
            verdict, vcol = "하락추세 — 조건 미충족", "red"

        bt = backtest(c, dts, st_dir, e200, adx)
        mc = None
        if bt and bt.get("durations"):
            mc = monte_carlo(c, bt["durations"], bt["total_return_pct"])
            bt.pop("durations", None)

        out_t.append({
            "symbol": sym, "name": name, "tag": tag,
            "price": round(c[-1], 2), "asof": asof,
            "trend_up": up, "st_long": st_long, "st_line": round(st_line[-1], 2),
            "adx": round(adx[-1], 1), "adx_strong": strong,
            "ema200": round(e200[-1], 2), "ema20": round(e20[-1], 2), "pullback": pullback,
            "verdict": verdict, "verdict_color": vcol,
            "backtest": bt, "monte_carlo": mc,
        })
        print(f"  {name}: {verdict} | ST {'롱' if st_long else '숏'} · ADX {adx[-1]:.0f} · "
              f"BT {bt['total_return_pct'] if bt else '—'}% (승률 {bt['win_rate'] if bt else '—'}%) · "
              f"MC 백분위 {mc['pctile'] if mc else '—'}")

    if not out_t:
        print("[ERROR] 전 종목 실패 — 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": asof,
        "params": {"supertrend": f"({ST_PERIOD},{ST_MULT})", "anchor": "EMA200", "adx_min": ADX_MIN,
                   "cost_pct": COST * 100},
        "tickers": out_t,
        "note": ("일봉 추세추종 신호(정보용·매매 실행 없음): 대추세 EMA200 + Supertrend(10,3) + ADX≥20, "
                 "풀백=상승추세 중 EMA20 아래 눌림. 백테스트 5년 롱온리·왕복비용 0.28%, "
                 "청산=Supertrend 트레일링. 몬테카를로=같은 횟수·보유기간 무작위 진입 1,000회 대비 백분위. "
                 "과거 성과≠미래. 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE} ({len(out_t)}종목 · 기준 {asof})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
