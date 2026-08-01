"""
추세추종 신호 보드 — Supertrend·EMA200·ADX + 다기간 백테스트 + 몬테카를로 (Jesse 리포트 스타일)
================================================================================================
알고트레이딩(Jesse) 방법론을 한국 시장 '일봉'에 적용한 정보 패널용 수집기.
  · 대추세: 종가 > EMA200 (앵커) / 신호: Supertrend(10,3) / 확인: ADX(14) ≥ 20
  · 백테스트(롱온리): 조건 충족 진입, Supertrend 하향 전환 청산(트레일링). 왕복비용 0.28%.
  · 다기간 리포트: YTD · 2.5년 · 5년 (Jesse 영상 워크플로 재현)
  · 월별 수익률 그리드 · 최악 낙폭 구간 Top3 · 전체 거래 목록(마커용)
  · 몬테카를로(랜덤진입 1,000회): 원본 샤프 vs null 중앙값/상위5% + 히스토그램
🚨 매매 실행 없음 — 신호 정보 제공용. 과거 성과≠미래. 투자자문 아님.
출력: docs/trend_signal.json
"""

import json
import math
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
    ("108490.KQ", "로보티즈", "보유"),
    ("005930.KS", "삼성전자", "관심"),
    ("003550.KS", "LG", "관심"),
    ("066570.KS", "LG전자", "관심"),
    ("042700.KS", "한미반도체", "관심"),
    ("373220.KS", "LG에너지솔루션", "관심"),
    ("247540.KQ", "에코프로비엠", "관심"),
]
ST_PERIOD, ST_MULT, ADX_MIN = 10, 3.0, 20
COST = 0.0028                         # 왕복 비용(수수료 0.05%×2 + 거래세 0.18%)
WARM = 201                            # EMA200 워밍업


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
            a = tr if a is None else (a * i + tr) / (i + 1)
        else:
            a = (a * (n - 1) + tr) / n
        out.append(a)
    return out


def supertrend(h, l, c, n=10, mult=3.0):
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
        for x in xs:
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
    for x in dx:
        a = x if a is None else (a * (n - 1) + x) / n
        out.append(a)
    return out


def run_strategy(c, dates, st_dir, e200, adx):
    """전 구간 1회 실행 → (일별 전략자산 curve_eq, 벤치마크 curve_bh, 날짜, 거래목록[인덱스 포함])."""
    n = len(c)
    trades, pos, entry_eq = [], None, 1.0
    eq = 1.0
    curve_eq, curve_bh, curve_d = [], [], []
    for i in range(WARM, n):
        if pos is None:
            if st_dir[i] == 1 and c[i] > e200[i] and adx[i] >= ADX_MIN:
                pos, entry_eq = i, eq
        else:
            eq *= c[i] / c[i - 1]
            if st_dir[i] == -1:
                eq *= (1 - COST)
                trades.append({"ei": pos, "xi": i, "entry_d": dates[pos], "exit_d": dates[i],
                               "entry_p": round(c[pos], 2), "exit_p": round(c[i], 2),
                               "ret": eq / entry_eq - 1, "days": i - pos})
                pos = None
        curve_eq.append(eq)
        curve_bh.append(c[i] / c[WARM])
        curve_d.append(dates[i])
    if pos is not None:
        trades.append({"ei": pos, "xi": n - 1, "entry_d": dates[pos], "exit_d": dates[n - 1],
                       "entry_p": round(c[pos], 2), "exit_p": round(c[n - 1], 2),
                       "ret": eq / entry_eq - 1, "days": n - 1 - pos, "open": True})
    return curve_eq, curve_bh, curve_d, trades


def _sharpe(dr):
    if len(dr) < 20:
        return None
    mu = sum(dr) / len(dr)
    sd = (sum((x - mu) ** 2 for x in dr) / (len(dr) - 1)) ** 0.5
    if not sd or sd != sd or mu != mu:                 # 0·NaN 방어
        return None
    return round(mu / sd * (252 ** 0.5), 2)


def period_report(curve_eq, curve_bh, curve_d, trades, start_ci):
    """curve 인덱스 start_ci부터 구간 리포트 (수익 재정규화·거래 필터)."""
    eqs = curve_eq[start_ci:]
    bhs = curve_bh[start_ci:]
    ds = curve_d[start_ci:]
    if len(eqs) < 30:
        return None
    e0, b0 = eqs[0], bhs[0]
    eqn = [v / e0 for v in eqs]
    bhn = [v / b0 for v in bhs]
    tr = [t for t in trades if t["xi"] - (len(curve_d) - len(ds)) >= 0 and t["exit_d"] >= ds[0]]
    tr_in = [t for t in tr if t["entry_d"] >= ds[0]]
    use = tr_in if tr_in else tr
    rets = [t["ret"] for t in use]
    wins = [r for r in rets if r > 0]
    losses = [r for r in rets if r <= 0]
    years = len(eqs) / 252
    dr = [eqn[i] / eqn[i - 1] - 1 for i in range(1, len(eqn))]
    peak, mdd, uw, max_uw = eqn[0], 0.0, 0, 0
    for v in eqn:
        if v >= peak:
            peak, uw = v, 0
        else:
            uw += 1
            max_uw = max(max_uw, uw)
            mdd = min(mdd, v / peak - 1)
    avg_win = sum(wins) / len(wins) * 100 if wins else None
    avg_loss = sum(losses) / len(losses) * 100 if losses else None
    step = max(1, len(eqn) // 200)
    samp = list(range(0, len(eqn), step))
    if samp[-1] != len(eqn) - 1:
        samp.append(len(eqn) - 1)
    return {
        "n_trades": len(use), "win_rate": round(len(wins) / len(use) * 100) if use else None,
        "avg_win_pct": round(avg_win, 2) if avg_win is not None else None,
        "avg_loss_pct": round(avg_loss, 2) if avg_loss is not None else None,
        "wl_ratio": round(avg_win / abs(avg_loss), 2) if (avg_win is not None and avg_loss) else None,
        "expectancy_pct": round(sum(rets) / len(rets) * 100, 2) if rets else None,
        "total_return_pct": round((eqn[-1] - 1) * 100, 1),
        "bench_return_pct": round((bhn[-1] - 1) * 100, 1),
        "cagr_pct": round(((eqn[-1] ** (1 / years)) - 1) * 100, 1) if years > 0.5 else None,
        "sharpe": _sharpe(dr), "mdd_pct": round(mdd * 100, 1), "max_underwater_days": max_uw,
        "trades_per_month": round(len(use) / (years * 12), 1) if years > 0.1 else None,
        "avg_hold_days": round(sum(t["days"] for t in use) / len(use)) if use else None,
        "curve": {"d0": ds[0], "d1": ds[-1],
                  "dd": [ds[i] for i in samp],
                  "eq": [round(eqn[i], 4) for i in samp],
                  "bh": [round(bhn[i], 4) for i in samp]},
    }


def monthly_table(curve_eq, curve_d):
    """월별 전략 수익률(%) — {연도: {월: pct}} + 연간 합계."""
    out = {}
    prev_eq, prev_ym = None, None
    month_start = curve_eq[0]
    for i, d in enumerate(curve_d):
        ym = d[:7]
        if prev_ym is not None and ym != prev_ym:
            y, m = prev_ym.split("-")
            out.setdefault(y, {})[int(m)] = round((prev_eq / month_start - 1) * 100, 1)
            month_start = prev_eq
        prev_eq, prev_ym = curve_eq[i], ym
    if prev_ym:
        y, m = prev_ym.split("-")
        out.setdefault(y, {})[int(m)] = round((prev_eq / month_start - 1) * 100, 1)
    yearly = {}
    for y in out:
        tot = 1.0
        for m in out[y]:
            tot *= (1 + out[y][m] / 100)
        yearly[y] = round((tot - 1) * 100, 1)
    return out, yearly


def worst_drawdowns(curve_eq, curve_d, top=3):
    """최악 낙폭 구간 Top N: (고점일, 저점일, 깊이%, 지속 거래일)."""
    eps = []
    peak, peak_i, trough, trough_i = curve_eq[0], 0, curve_eq[0], 0
    in_dd = False
    for i, v in enumerate(curve_eq):
        if v >= peak:
            if in_dd:
                eps.append({"from": curve_d[peak_i], "trough": curve_d[trough_i],
                            "depth_pct": round((trough / peak - 1) * 100, 1), "days": i - peak_i})
                in_dd = False
            peak, peak_i = v, i
            trough, trough_i = v, i
        else:
            in_dd = True
            if v < trough:
                trough, trough_i = v, i
    if in_dd:
        eps.append({"from": curve_d[peak_i], "trough": curve_d[trough_i],
                    "depth_pct": round((trough / peak - 1) * 100, 1), "days": len(curve_eq) - 1 - peak_i,
                    "ongoing": True})
    eps.sort(key=lambda e: e["depth_pct"])
    return eps[:top]


def run_strategy_short(c, dates, st_dir, e200, adx):
    """숏 사이드 시뮬레이션(정보용): ST 숏 + 종가<EMA200 + ADX≥20 진입, ST 상향 전환 커버.
    ※ 한국 개인은 공매도 제약이 커서 '실행'보다는 전략 대칭성 확인용. 일수익 근사 = -(일간변동)."""
    n = len(c)
    trades, pos, entry_eq = [], None, 1.0
    eq = 1.0
    for i in range(WARM, n):
        if pos is None:
            if st_dir[i] == -1 and c[i] < e200[i] and adx[i] >= ADX_MIN:
                pos, entry_eq = i, eq
        else:
            eq *= max(0.01, 2 - c[i] / c[i - 1])          # 숏 일수익 = 1 - 일간등락
            if st_dir[i] == 1:
                eq *= (1 - COST)
                trades.append({"entry_d": dates[pos], "exit_d": dates[i],
                               "entry_p": round(c[pos], 2), "exit_p": round(c[i], 2),
                               "ei": pos, "xi": i,
                               "ret": eq / entry_eq - 1, "days": i - pos})
                pos = None
    if pos is not None:
        trades.append({"entry_d": dates[pos], "exit_d": dates[n - 1],
                       "entry_p": round(c[pos], 2), "exit_p": round(c[n - 1], 2),
                       "ei": pos, "xi": n - 1,
                       "ret": eq / entry_eq - 1, "days": n - 1 - pos, "open": True})
    return trades


def run_strategy_limit(c, l, dates, st_dir, e200, e20, adx):
    """리밋주문 풀백 진입 변형: 조건 충족 시 '대기' → 전일 EMA20에 리밋 걸고
    저가가 닿으면 체결(진입가=전일 EMA20). 청산은 동일(ST 하향 트레일링).
    영상에서 '풀백 진입은 원래 리밋주문 의도였다'고 지적한 부분의 구현."""
    n = len(c)
    trades, pos, entry_eq = [], None, 1.0
    eq = 1.0
    curve_eq = []
    for i in range(WARM, n):
        if pos is None:
            armed = st_dir[i] == 1 and c[i] > e200[i] and adx[i] >= ADX_MIN
            # 리밋 체결 판정: 조건 충족 대기 상태에서 저가가 전일 EMA20 이하로 눌리면 체결
            if armed and l[i] <= e20[i - 1] and c[i] > 0:
                fill = min(e20[i - 1], c[i - 1])          # 갭하락 보수 반영
                eq *= c[i] / fill
                pos, entry_eq = i, eq / (c[i] / fill)
        else:
            eq *= c[i] / c[i - 1]
            if st_dir[i] == -1:
                eq *= (1 - COST)
                trades.append({"ret": eq / entry_eq - 1, "days": i - pos})
                pos = None
        curve_eq.append(eq)
    if pos is not None:
        trades.append({"ret": eq / entry_eq - 1, "days": n - 1 - pos, "open": True})
    return curve_eq, trades


def quick_metrics(curve_eq, trades):
    """(총수익%, 샤프, MDD%, 승률, 거래수) — 최적화·비교표용 경량 지표."""
    if not trades or len(curve_eq) < 30:
        return None
    dr = [curve_eq[i] / curve_eq[i - 1] - 1 for i in range(1, len(curve_eq))]
    peak, mdd = curve_eq[0], 0.0
    for v in curve_eq:
        peak = max(peak, v)
        mdd = min(mdd, v / peak - 1)
    closed = [t for t in trades]
    wins = sum(1 for t in closed if t["ret"] > 0)
    return {"total_return_pct": round((curve_eq[-1] - 1) * 100, 1), "sharpe": _sharpe(dr),
            "mdd_pct": round(mdd * 100, 1), "win_rate": round(wins / len(closed) * 100),
            "n_trades": len(closed)}


def run_variant(c, dates, st_dir, e200, adx, mode):
    """구성요소 변형 백테스트 — '엣지가 어디서 오는가' 분해용.
    영상 전략 주석의 WHERE THE EDGE COMES FROM 접근: 규칙을 하나씩 떼어내 각각 측정."""
    n = len(c)
    eq, pos, entry_eq, trades, curve = 1.0, None, 1.0, [], []
    for i in range(WARM, n):
        if mode == "trend":                       # 대추세만 (EMA200 위/아래)
            enter, exit_ = c[i] > e200[i], c[i] < e200[i]
        elif mode == "st":                        # Supertrend만
            enter, exit_ = st_dir[i] == 1, st_dir[i] == -1
        elif mode == "st_trend":                  # ST + 대추세 (ADX 없음)
            enter, exit_ = (st_dir[i] == 1 and c[i] > e200[i]), st_dir[i] == -1
        elif mode == "full_adx30":                # 전체 + ADX 강화(≥30)
            enter, exit_ = (st_dir[i] == 1 and c[i] > e200[i] and adx[i] >= 30), st_dir[i] == -1
        else:                                     # full = 현재 규칙 (ADX≥20)
            enter, exit_ = (st_dir[i] == 1 and c[i] > e200[i] and adx[i] >= ADX_MIN), st_dir[i] == -1
        if pos is None:
            if enter:
                pos, entry_eq = i, eq
        else:
            eq *= c[i] / c[i - 1]
            if exit_:
                eq *= (1 - COST)
                trades.append({"ret": eq / entry_eq - 1, "days": i - pos})
                pos = None
        curve.append(eq)
    if pos is not None:
        trades.append({"ret": eq / entry_eq - 1, "days": n - 1 - pos, "open": True})
    return quick_metrics(curve, trades)


def edge_attribution(c, dates, st_dir, e200, adx):
    """엣지 귀속 분석 — 어느 구성요소가 실제 수익/방어를 만드는가."""
    SPECS = [("trend", "① 대추세만 (종가>EMA200)"),
             ("st", "② Supertrend만 (10,3)"),
             ("st_trend", "③ ST + 대추세 (ADX 없음)"),
             ("full", "④ 전체 = ③ + ADX≥20  ← 현재"),
             ("full_adx30", "⑤ 전체 + ADX≥30 (강화)")]
    rows = []
    for mode, label in SPECS:
        q = run_variant(c, dates, st_dir, e200, adx, mode)
        if q:
            rows.append({"mode": mode, "label": label, "current": mode == "full", **q})
    if not rows:
        return None
    full = next((r for r in rows if r["mode"] == "full"), None)
    note = ""
    if full:
        best_ret = max(rows, key=lambda r: r["total_return_pct"])
        best_mdd = min(rows, key=lambda r: r["mdd_pct"])          # 값이 클수록(0에 가까울수록) 방어 우수
        best_mdd = max(rows, key=lambda r: r["mdd_pct"])
        parts = []
        if best_ret["mode"] != "full":
            parts.append(f"수익은 '{best_ret['label'].split('(')[0].strip()}'가 더 높음({best_ret['total_return_pct']:+.1f}% vs {full['total_return_pct']:+.1f}%)")
        else:
            parts.append("현재 조합이 수익 최고")
        if best_mdd["mode"] != "full":
            parts.append(f"낙폭 방어는 '{best_mdd['label'].split('(')[0].strip()}'가 우수(MDD {best_mdd['mdd_pct']}% vs {full['mdd_pct']}%)")
        else:
            parts.append("현재 조합이 낙폭 방어 최고")
        st_only = next((r for r in rows if r["mode"] == "st"), None)
        tr_only = next((r for r in rows if r["mode"] == "trend"), None)
        if st_only and tr_only:
            parts.append("엣지 원천: " + ("대추세 필터" if tr_only["total_return_pct"] > st_only["total_return_pct"] else "Supertrend 신호"))
        note = " · ".join(parts)
    return {"rows": rows, "note": note}


def optimize_grid(h, l, c, dates, e200, adx):
    """Supertrend 기간×배수 그리드 탐색 (5년 전체) — 샤프 내림차순 상위."""
    grid = []
    for p in (7, 10, 14):
        for m in (2.0, 2.5, 3.0, 3.5, 4.0):
            sd, _ = supertrend(h, l, c, p, m)
            ce, _, _, tr = run_strategy(c, dates, sd, e200, adx)
            qm = quick_metrics(ce, tr)
            if qm:
                qm.update({"period": p, "mult": m,
                           "current": (p == ST_PERIOD and abs(m - ST_MULT) < 0.01)})
                grid.append(qm)
    grid.sort(key=lambda g: (g["sharpe"] if g["sharpe"] is not None else -9), reverse=True)
    return grid


def rule_significance(c, st_dir, e200, adx, sims=2000, seed=11):
    """RST(Rule Significance Test) — 진입 규칙의 '봉 단위' 유의성 검정.
    귀무가설 H₀: 규칙이 고른 시점은 무작위 시점과 다르지 않다.
      · 추세 편향 제거를 위해 전체 로그수익률을 디트렌드(평균 차감) → 상승장 프리미엄 상쇄
      · 규칙이 '보유 중'이던 봉들의 평균 디트렌드 수익률 = 관측값
      · 같은 개수의 봉을 무작위 추출해 평균을 2,000회 계산 → H₀ 분포
      · p = (관측값 이상인 시뮬레이션 비율) — p≤0.05면 '규칙에 엣지 있음'
    ※ 몬테카를로(거래 단위 타이밍 검정)와 보완 관계. 진입규칙 자체의 노이즈 여부를 봄."""
    n = len(c)
    if n <= WARM + 30:
        return None
    lr = [math.log(c[i] / c[i - 1]) for i in range(WARM, n) if c[i - 1] > 0]
    if len(lr) < 50:
        return None
    mu_all = sum(lr) / len(lr)
    det = [x - mu_all for x in lr]                      # 디트렌드: 시장 드리프트 제거

    in_pos, pos = [], False                              # 규칙 보유 여부(각 봉)
    for i in range(WARM, n):
        if not pos:
            pos = (st_dir[i] == 1 and c[i] > e200[i] and adx[i] >= ADX_MIN)
            in_pos.append(False)                         # 진입 당일은 다음 봉부터 수익 귀속
        else:
            in_pos.append(True)
            if st_dir[i] == -1:
                pos = False
    m = min(len(in_pos), len(det))
    sel = [det[i] for i in range(m) if in_pos[i]]
    if len(sel) < 20:
        return None
    obs = sum(sel) / len(sel)

    rng = random.Random(seed)
    k, pool = len(sel), det
    null = []
    for _ in range(sims):
        s = 0.0
        for _ in range(k):
            s += pool[rng.randrange(len(pool))]
        null.append(s / k)
    null.sort()
    ge = sum(1 for x in null if x >= obs)
    p = round(ge / sims, 4)
    lo, hi = min(null[0], obs), max(null[-1], obs)
    span = (hi - lo) or 1e-12
    NB = 40
    bins = [0] * NB
    for x in null:
        bins[min(NB - 1, int((x - lo) / span * NB))] += 1
    obs_bin = min(NB - 1, max(0, int((obs - lo) / span * NB)))
    ann = (math.exp(obs * 252) - 1) * 100                # 일봉 기준 연환산(디트렌드 초과분)
    return {"observed_mean": obs, "p_value": p, "significant": p <= 0.05,
            "n_bars": k, "n_pool": len(pool), "sims": sims,
            "annualised_excess_pct": round(ann, 2),
            "hist": {"bins": bins, "lo": lo, "hi": hi, "obs_bin": obs_bin},
            "verdict": ("통계적으로 유의 ✅ (규칙에 엣지)" if p <= 0.05
                        else "약한 신호 (관찰 지속)" if p <= 0.15 else "노이즈와 구분 불가"),
            "verdict_color": "green" if p <= 0.05 else "yellow" if p <= 0.15 else "red"}


def mc_candles(h, l, c, dates, orig, scen=200, block=10, seed=23, keep=40):
    """Monte Carlo candles — '가격 경로'를 블록 부트스트랩으로 재생성해 전략을 매번 재실행.
    같은 통계적 성질(변동성·자기상관)의 다른 시나리오 200개에서도 성과가 유지되는지 = 강건성 검정.
      · 블록 단위(기본 10봉) 재표본 → 추세·군집 변동성 보존 (단순 IID 셔플의 한계 보완)
      · 각 시나리오마다 EMA200·Supertrend·ADX 재계산 후 동일 규칙으로 백테스트
      · 결과: 원본 vs 하위5%·중앙값·상위5% 지표표 + 자산곡선 샘플(스파게티용)"""
    n = len(c)
    if n <= WARM + 60 or not orig:
        return None
    rets = [c[i] / c[i - 1] for i in range(1, n) if c[i - 1] > 0]
    hr = [h[i] / c[i] for i in range(1, n) if c[i] > 0]              # 고가·저가 비율 보존
    lr_ = [l[i] / c[i] for i in range(1, n) if c[i] > 0]
    m = min(len(rets), len(hr), len(lr_))
    if m < 100:
        return None
    rng = random.Random(seed)
    nb = m // block + 1
    curves, mets = [], {"total": [], "sharpe": [], "mdd": [], "win": [], "n": []}
    for _ in range(scen):
        idx = []
        for _ in range(nb):                                          # 블록 부트스트랩
            s = rng.randrange(0, max(1, m - block))
            idx.extend(range(s, min(m, s + block)))
        idx = idx[:m]
        sc_, sh_, sl_ = [c[0]], [h[0]], [l[0]]
        for j in idx:
            p = sc_[-1] * rets[j]
            sc_.append(p)
            sh_.append(p * hr[j])
            sl_.append(p * lr_[j])
        e2 = ema(sc_, 200)
        sd, _ln = supertrend(sh_, sl_, sc_, ST_PERIOD, ST_MULT)
        ax = adx_wilder(sh_, sl_, sc_, 14)
        ce, _bh, _cd, tr = run_strategy(sc_, [""] * len(sc_), sd, e2, ax)
        q = quick_metrics(ce, tr)
        if not q:
            continue
        mets["total"].append(q["total_return_pct"]); mets["mdd"].append(q["mdd_pct"])
        mets["win"].append(q["win_rate"]); mets["n"].append(q["n_trades"])
        if q["sharpe"] is not None:
            mets["sharpe"].append(q["sharpe"])
        if len(curves) < keep and len(ce) > 10:                      # 표시용 곡선만 보관
            st_ = max(1, len(ce) // 40)
            curves.append([round(ce[i], 3) for i in range(0, len(ce), st_)][:41])
    if not mets["total"]:
        return None

    def pct_(a, q):
        if not a:
            return None
        b = sorted(a)
        return round(b[min(len(b) - 1, max(0, int(len(b) * q)))], 2)
    rows = []
    for key, lab, o in (("total", "총 순수익", orig.get("total_return_pct")),
                        ("mdd", "최대 낙폭(MDD)", orig.get("mdd_pct")),
                        ("sharpe", "샤프 지수", orig.get("sharpe")),
                        ("win", "승률", orig.get("win_rate")),
                        ("n", "총 거래수", orig.get("n_trades"))):
        a = mets[key]
        rows.append({"metric": lab, "original": o,
                     "worst5": pct_(a, 0.05), "median": pct_(a, 0.5), "best5": pct_(a, 0.95)})
    tot = sorted(mets["total"])
    below = sum(1 for x in tot if x < (orig.get("total_return_pct") or 0))
    pctile = round(below / len(tot) * 100, 1)
    if pctile <= 50:
        v, col = "원본이 중앙값 이하 — 과최적화 징후 낮음 ✅", "green"
    elif pctile <= 80:
        v, col = "원본이 상위권 — 운이 일부 작용 가능", "yellow"
    else:
        v, col = "원본이 상위 20% — 과최적화 의심 ⚠️", "red"
    return {"scenarios": len(mets["total"]), "block": block, "rows": rows,
            "curves": curves, "orig_pctile": pctile, "verdict": v, "verdict_color": col}


def monte_carlo(c, trades, orig_sharpe, sims=1000, seed=7):
    """랜덤진입 검정(샤프 기준): 같은 횟수·보유기간 무작위 진입 null 분포 → 원본 위치.
    Jesse 영상 해석: 원본이 null 중앙값 근처/이하 = 과적합 아님(운 아님)이 아니라,
    '무작위 대비 우위'는 원본이 상위(95↑)일 때. 두 관점 모두 표기용 수치 제공."""
    durations = [t["days"] for t in trades if t["days"] > 0]
    if not durations or orig_sharpe is None:
        return None
    rng = random.Random(seed)
    n = len(c)
    null_sharpes, null_totals = [], []
    for _ in range(sims):
        dr, tot = [], 1.0
        for d in durations:
            s = rng.randrange(WARM, n - d) if n - d > WARM else WARM
            for j in range(s + 1, s + d + 1):
                dr.append(c[j] / c[j - 1] - 1)
            tot *= (1 + (c[s + d] / c[s] - 1 - COST))
        sh = _sharpe(dr)
        if sh is not None:
            null_sharpes.append(sh)
        null_totals.append((tot - 1) * 100)
    if not null_sharpes:
        return None
    null_sharpes.sort()
    ns = len(null_sharpes)
    med = null_sharpes[ns // 2]
    best5 = null_sharpes[int(ns * 0.95)]
    below = sum(1 for x in null_sharpes if x < orig_sharpe)
    pctile = round(below / ns * 100, 1)
    # 히스토그램 (null 샤프 20구간)
    lo, hi = null_sharpes[0], null_sharpes[-1]
    if orig_sharpe < lo:
        lo = orig_sharpe
    if orig_sharpe > hi:
        hi = orig_sharpe
    span = (hi - lo) or 1
    bins = [0] * 20
    for x in null_sharpes:
        bins[min(19, int((x - lo) / span * 20))] += 1
    orig_bin = min(19, max(0, int((orig_sharpe - lo) / span * 20)))
    if pctile >= 95:
        v, col = "무작위 대비 유의미한 엣지 ✅", "green"
    elif pctile >= 80:
        v, col = "무작위보다 우수 (경계선)", "yellow"
    else:
        v, col = "무작위 진입과 구분 불가", "red"
    return {"sims": ns, "orig_sharpe": orig_sharpe, "null_median_sharpe": med,
            "null_best5_sharpe": best5, "pctile": pctile,
            "hist": {"bins": bins, "lo": round(lo, 2), "hi": round(hi, 2), "orig_bin": orig_bin},
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
        if df is None or len(df) < 300:
            print(f"  [WARN] {name} 데이터 부족")
            continue
        df = df.dropna(subset=["High", "Low", "Close"])          # NaN 행 제거 (야후 데이터 결측 방어)
        if len(df) < 300:
            print(f"  [WARN] {name} 유효 데이터 부족")
            continue
        h, l, c = df["High"].tolist(), df["Low"].tolist(), df["Close"].tolist()
        dts = [x.strftime("%Y-%m-%d") for x in df.index]
        e200, e20 = ema(c, 200), ema(c, 20)
        st_dir, st_line = supertrend(h, l, c, ST_PERIOD, ST_MULT)
        adx = adx_wilder(h, l, c, 14)
        asof = dts[-1]

        # 현재 신호
        up, st_long, strong = c[-1] > e200[-1], st_dir[-1] == 1, adx[-1] >= ADX_MIN
        pullback = up and st_long and c[-1] < e20[-1]
        if up and st_long and strong:
            verdict, vcol = ("추세추종 조건 충족 (풀백 중)" if pullback else "추세추종 조건 충족"), "green"
        elif up and st_long:
            verdict, vcol = "상승추세 · 추세강도 약함(ADX<20)", "yellow"
        elif up:
            verdict, vcol = "대추세 상승 · Supertrend 하향(조정 중)", "yellow"
        else:
            verdict, vcol = "하락추세 — 조건 미충족", "red"

        curve_eq, curve_bh, curve_d, trades = run_strategy(c, dts, st_dir, e200, adx)
        if not trades:
            print(f"  [WARN] {name} 거래 없음")
            continue

        # 다기간 리포트: 5년(전체) · 2.5년 · YTD
        year0 = now.strftime("%Y") + "-01-01"
        idx_25 = max(0, len(curve_d) - int(2.5 * 252))
        idx_ytd = next((i for i, d in enumerate(curve_d) if d >= year0), None)
        periods = {"5y": period_report(curve_eq, curve_bh, curve_d, trades, 0),
                   "2.5y": period_report(curve_eq, curve_bh, curve_d, trades, idx_25)}
        if idx_ytd is not None and len(curve_d) - idx_ytd >= 30:
            periods["ytd"] = period_report(curve_eq, curve_bh, curve_d, trades, idx_ytd)

        monthly, yearly = monthly_table(curve_eq, curve_d)
        wdd = worst_drawdowns(curve_eq, curve_d, 3)
        p5 = periods.get("5y") or {}
        mc = monte_carlo(c, trades, p5.get("sharpe"))
        rst = rule_significance(c, st_dir, e200, adx)                 # 봉 단위 규칙 유의성
        mcc = mc_candles(h, l, c, dts, p5)                            # 가격경로 재표본 강건성
        attr = edge_attribution(c, dts, st_dir, e200, adx)            # 엣지 귀속(구성요소 기여)

        # 하이퍼파라미터 그리드 탐색 (Supertrend 기간×배수, 15조합)
        opt = optimize_grid(h, l, c, dts, e200, adx)
        # 진입방식 비교: 시장가(조건 충족 즉시) vs 리밋(전일 EMA20 풀백 체결)
        ce_l, tr_l = run_strategy_limit(c, l, dts, st_dir, e200, e20, adx)
        entry_compare = {"market": quick_metrics(curve_eq, trades),
                         "limit": quick_metrics(ce_l, tr_l)}

        # 거래 목록 (Jesse Trade chart용 — 전체, 진입/청산가 포함, 최대 40건)
        tlist = [{"entry_d": t["entry_d"], "exit_d": t["exit_d"],
                  "entry_p": t["entry_p"], "exit_p": t["exit_p"],
                  "ret_pct": round(t["ret"] * 100, 1), "days": t["days"],
                  "open": bool(t.get("open"))} for t in trades][-40:]
        # 트레이드 차트 오버레이: EMA200·Supertrend 라인 (5y curve와 동일 샘플링·정규화)
        L = len(curve_d)
        _step = max(1, L // 200)
        _samp = list(range(0, L, _step))
        if _samp[-1] != L - 1:
            _samp.append(L - 1)
        _base = c[WARM]
        overlay = {"e200": [round(e200[WARM + i] / _base, 4) for i in _samp],
                   "st": [round(st_line[WARM + i] / _base, 4) for i in _samp],
                   "adx": [round(adx[WARM + i], 1) for i in _samp]}
        # 거래 클릭 줌용 일별 시계열 (실제 가격 — 정수 반올림으로 용량 절감)
        daily = {"dates": dts[WARM:],
                 "c": [round(x) for x in c[WARM:]],
                 "e200": [round(x) for x in e200[WARM:]],
                 "st": [round(x) for x in st_line[WARM:]]}

        # 숏 사이드 시뮬레이션 (정보용 — 영상의 SHORT 마커 재현)
        st_trades = run_strategy_short(c, dts, st_dir, e200, adx)
        d2i_s = {dd: ii for ii, dd in enumerate(curve_d)}
        nn_s = len(curve_d) - 1
        tlist_short = [{"entry_d": x["entry_d"], "exit_d": x["exit_d"],
                        "entry_p": x["entry_p"], "exit_p": x["exit_p"],
                        "ret_pct": round(x["ret"] * 100, 1), "days": x["days"],
                        "open": bool(x.get("open"))} for x in st_trades][-40:]
        markers_short = [{"e": round(d2i_s.get(x["entry_d"], 0) / nn_s, 4),
                          "x": round(d2i_s.get(x["exit_d"], 0) / nn_s, 4),
                          "win": x["ret"] > 0} for x in st_trades]
        s_tot = 1.0
        for x in st_trades:
            s_tot *= (1 + x["ret"])
        s_wins = sum(1 for x in st_trades if x["ret"] > 0)
        short_stats = {"n": len(st_trades), "wins": s_wins,
                       "win_rate": round(s_wins / len(st_trades) * 100) if st_trades else None,
                       "total_return_pct": round((s_tot - 1) * 100, 1) if st_trades else None}
        # 가격(벤치마크) 곡선 위 거래 마커 좌표 (곡선 분율 0~1)
        d2i = {d: i for i, d in enumerate(curve_d)}
        nn = len(curve_d) - 1
        markers = [{"e": round(d2i.get(t["entry_d"], 0) / nn, 4), "x": round(d2i.get(t["exit_d"], 0) / nn, 4),
                    "win": t["ret"] > 0} for t in trades]

        out_t.append({
            "symbol": sym, "name": name, "tag": tag,
            "price": round(c[-1], 2), "asof": asof,
            "trend_up": up, "st_long": st_long, "st_line": round(st_line[-1], 2),
            "adx": round(adx[-1], 1), "adx_strong": strong,
            "ema200": round(e200[-1], 2), "ema20": round(e20[-1], 2), "pullback": pullback,
            "verdict": verdict, "verdict_color": vcol,
            "periods": periods, "monthly": monthly, "yearly": yearly,
            "worst_dd": wdd, "trades": tlist, "markers": markers, "overlay": overlay, "daily": daily,
            "trades_short": tlist_short, "markers_short": markers_short, "short_stats": short_stats,
            "monte_carlo": mc, "rst": rst, "mc_candles": mcc, "attribution": attr,
            "optimization": opt[:8], "entry_compare": entry_compare,
        })
        print(f"  {name}: {verdict}")
        for pk in ("ytd", "2.5y", "5y"):
            p = periods.get(pk)
            if p:
                print(f"    [{pk:4s}] 수익 {p['total_return_pct']:+7.1f}% (벤치 {p['bench_return_pct']:+8.1f}%) "
                      f"샤프 {p['sharpe']} 승률 {p['win_rate']}% MDD {p['mdd_pct']}% 수중 {p['max_underwater_days']}일")
        if mc:
            print(f"    MC: 원본샤프 {mc['orig_sharpe']} vs null중앙 {mc['null_median_sharpe']} (백분위 {mc['pctile']}) → {mc['verdict']}")

    if not out_t:
        print("[ERROR] 전 종목 실패 — 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": asof,
        "params": {"supertrend": f"({ST_PERIOD},{ST_MULT})", "anchor": "EMA200", "adx_min": ADX_MIN,
                   "cost_pct": COST * 100},
        "tickers": out_t,
        "note": ("일봉 추세추종 신호(정보용·매매 실행 없음): 대추세 EMA200 + Supertrend(10,3) + ADX≥20. "
                 "백테스트 롱온리·왕복비용 0.28%·청산=Supertrend 트레일링. 다기간(YTD·2.5y·5y) 리포트, "
                 "월별 수익률, 최악 낙폭 구간, 몬테카를로(랜덤진입 1,000회·샤프 기준). "
                 "과거 성과≠미래 · 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE} ({len(out_t)}종목 · 기준 {asof} · {os.path.getsize(OUTPUT_FILE)//1024}KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
