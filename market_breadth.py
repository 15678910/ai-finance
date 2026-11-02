"""
시장폭(Market Breadth) 게이트 — 한국 시장 (시총 상위 종목 기준)
==================================================================
"시장 환경이 나쁘면 종목을 고르지 않는다"를 숫자로 만든다.

측정 (유니버스: 코스피·코스닥 시총 상위 각 400종목, ETF 제외):
  • 20·50·200일 이동평균 위에 있는 종목 비율
  • 52주 신고가 vs 신저가 종목 수, 당일 상승 vs 하락 종목 수
  • 종합 점수 0~100 = 평균(50일선 위 %, 200일선 위 %, 20일선 위 %, 신고가/(신고가+신저가) %)
    → 70+ 강세 · 50~70 중립 · 30~50 약세 · 30 미만 극약세
  • 게이트: 점수 30 미만 또는 50일선 위 비율 25% 미만이면 '차단'
지수 맥락 (코스피·코스닥):
  • 52주 고점 대비 하락률, 저점 대비 반등률·경과일
  • 팔로스루데이(FTD): 저점 후 4~12일차에 +1.5% 이상 & 거래량 전일 초과 (거래량은 KODEX 200 / KOSDAQ 150 ETF 대용)
  • 분배일: 최근 25거래일 중 -0.2% 이상 하락 & 거래량 전일 초과 일수 (6일 이상이면 천장 경고)

데이터: yfinance 일봉 1년. 수집률 50% 미만이면 결과를 덮어쓰지 않고 실패(exit 1)한다 — 조용히 틀린 값을 내지 않기 위해.
출력: docs/market_breadth.json  /  상태(점수 이력): docs/market_breadth_state.json
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
from core.state_store import load_state, save_state
from core.kr_universe import fetch_universe, yf_symbol

KST = timezone(timedelta(hours=9))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "market_breadth.json")
STATE_NAME = "market_breadth"
TOP_KOSPI, TOP_KOSDAQ = 400, 400
MIN_COVERAGE = 0.5
INDEX_PROXY = {"KOSPI": ("^KS11", "069500.KS"), "KOSDAQ": ("^KQ11", "229200.KS")}   # (지수, 거래량 대용 ETF)
MAX_HISTORY = 400
DIST_DAY_WARN = 6


# ====================================================================
# 순수 계산 (테스트 대상)
# ====================================================================
def sma(vals, n):
    if len(vals) < n:
        return None
    return sum(vals[-n:]) / n


def stock_metrics(closes):
    """한 종목의 종가 리스트(오래된→최신) → 지표 플래그. 이력이 모자란 지표는 None."""
    if not closes:
        return {}
    c = closes[-1]
    m = {}
    for n in (20, 50, 200):
        s = sma(closes, n)
        m[f"above{n}"] = (c > s) if s is not None else None
    if len(closes) >= 120:
        win = closes[-252:]
        m["high52"] = c >= max(win)
        m["low52"] = c <= min(win)
    else:
        m["high52"] = m["low52"] = None
    m["adv"] = (c > closes[-2]) if len(closes) >= 2 else None
    m["dec"] = (c < closes[-2]) if len(closes) >= 2 else None
    return m


def ma_distances(closes):
    """종가의 20/50/200일 이동평균 대비 거리(%). 이력 부족이면 None."""
    out = {}
    if not closes:
        return {"d20": None, "d50": None, "d200": None}
    c = closes[-1]
    for n in (20, 50, 200):
        s = sma(closes, n)
        out[f"d{n}"] = round((c / s - 1) * 100, 2) if s else None
    return out


def compute_breadth(hist):
    """{심볼: closes} → 집계. 비율은 해당 지표를 계산할 수 있었던 종목만 분모로 쓴다."""
    keys = ("above20", "above50", "above200", "high52", "low52", "adv", "dec")
    cnt = {k: 0 for k in keys}
    elig = {k: 0 for k in keys}
    n = 0
    for closes in hist.values():
        m = stock_metrics(closes)
        if not m:
            continue
        n += 1
        for k in keys:
            if m.get(k) is not None:
                elig[k] += 1
                cnt[k] += 1 if m[k] else 0
    pct = lambda k: round(cnt[k] / elig[k] * 100, 1) if elig[k] else None
    return {"n": n,
            "pct_above_20": pct("above20"), "pct_above_50": pct("above50"), "pct_above_200": pct("above200"),
            "eligible_200": elig["above200"],
            "new_highs": cnt["high52"], "new_lows": cnt["low52"],
            "advancers": cnt["adv"], "decliners": cnt["dec"]}


def score_breadth(b):
    """종합 점수·등급·게이트. 필요한 비율이 없으면 None 점수 + 게이트 '판단 불가'."""
    parts = [b.get("pct_above_50"), b.get("pct_above_200"), b.get("pct_above_20")]
    if any(p is None for p in parts):
        return {"score": None, "grade": "판단 불가", "gate_open": None, "reason": "이동평균 비율 계산 불가"}
    hi, lo = b.get("new_highs") or 0, b.get("new_lows") or 0
    hl = hi / (hi + lo) * 100 if (hi + lo) else 50.0
    score = round((parts[0] + parts[1] + parts[2] + hl) / 4, 1)
    grade = "강세" if score >= 70 else "중립" if score >= 50 else "약세" if score >= 30 else "극약세"
    reasons = []
    if score < 30:
        reasons.append(f"종합 {score}점 < 30")
    if parts[0] < 25:
        reasons.append(f"50일선 위 {parts[0]}% < 25%")
    gate_open = not reasons
    return {"score": score, "grade": grade, "gate_open": gate_open,
            "reason": " · ".join(reasons) if reasons else f"종합 {score}점 · 50일선 위 {parts[0]}%",
            "hl_pct": round(hl, 1)}


def index_context(closes, volumes, dates):
    """지수 맥락: 고점·저점·반등·FTD·분배일. closes/volumes/dates 는 같은 길이(오래된→최신)."""
    if not closes or len(closes) < 30:
        return {}
    win = closes[-252:]
    off = len(closes) - len(win)
    i_hi = off + max(range(len(win)), key=lambda i: win[i])
    lo_seg = closes[i_hi:]
    i_lo = i_hi + min(range(len(lo_seg)), key=lambda i: lo_seg[i])
    c, hi, lo = closes[-1], closes[i_hi], closes[i_lo]
    out = {"close": round(c, 2), "high_52w": round(hi, 2), "high_date": dates[i_hi],
           "drawdown_from_high_pct": round((c / hi - 1) * 100, 2),
           "low_since_high": round(lo, 2), "low_date": dates[i_lo],
           "max_drawdown_pct": round((lo / hi - 1) * 100, 2),
           "rebound_from_low_pct": round((c / lo - 1) * 100, 2),
           "days_since_low": len(closes) - 1 - i_lo}
    # FTD: 저점 후 4~12일차, +1.5% 이상 & 거래량 전일 초과
    ftd = None
    if volumes and len(volumes) == len(closes):
        for d in range(i_lo + 4, min(i_lo + 13, len(closes))):
            if closes[d] / closes[d - 1] - 1 >= 0.015 and volumes[d] > volumes[d - 1] > 0:
                ftd = {"date": dates[d], "day_n": d - i_lo, "gain_pct": round((closes[d] / closes[d - 1] - 1) * 100, 2)}
                break
        # 분배일: 최근 25거래일, -0.2% 이상 하락 & 거래량 증가
        dd = 0
        for d in range(max(1, len(closes) - 25), len(closes)):
            if closes[d] < closes[d - 1] * 0.998 and volumes[d] > volumes[d - 1] > 0:
                dd += 1
        out["distribution_days_25"] = dd
    out["ftd"] = ftd
    out["ftd_window_open"] = 4 <= out["days_since_low"] <= 12 and ftd is None
    s50, s200 = sma(closes, 50), sma(closes, 200)
    out["above_50dma"] = (c > s50) if s50 else None
    out["above_200dma"] = (c > s200) if s200 else None
    return out


# ====================================================================
# 수집
# ====================================================================
def fetch_price_history(symbols, period="1y", chunk=150):
    """{심볼: {"dates": [...], "close": [...], "volume": [...]}} — yfinance 일괄 다운로드."""
    import yfinance as yf
    out = {}
    for i in range(0, len(symbols), chunk):
        batch = symbols[i:i + chunk]
        try:
            df = yf.download(batch, period=period, interval="1d", group_by="ticker",
                             auto_adjust=False, threads=True, progress=False)
        except Exception as e:
            print(f"  [WARN] yf.download batch {i // chunk}: {type(e).__name__}: {e}")
            continue
        if df is None or df.empty:
            continue
        for sym in batch:
            try:
                sub = df[sym] if len(batch) > 1 else df
                sub = sub.dropna(subset=["Close"])
                if len(sub) < 2:
                    continue
                out[sym] = {"dates": [d.strftime("%Y-%m-%d") for d in sub.index],
                            "close": [float(x) for x in sub["Close"]],
                            "volume": [float(x) if x == x else 0.0 for x in sub["Volume"]]}
            except Exception:
                continue
    return out


def main():
    now = datetime.now(KST)
    print("=" * 60)
    print("  시장폭 게이트 (한국 시장 · 시총 상위 종목)")
    print(f"  KST: {now:%Y-%m-%d %H:%M:%S}")
    print("=" * 60)
    universe = fetch_universe(TOP_KOSPI, TOP_KOSDAQ)
    if not universe:
        print("  [FAIL] 유니버스 없음")
        return 1
    symbols = [yf_symbol(u["code"], u["market"]) for u in universe]
    market_of = {yf_symbol(u["code"], u["market"]): u["market"] for u in universe}
    hist = fetch_price_history(symbols)
    cov = len(hist) / len(symbols)
    print(f"  일봉 수집: {len(hist)}/{len(symbols)} ({cov:.0%})")
    if cov < MIN_COVERAGE:
        print(f"  [FAIL] 수집률 {cov:.0%} < {MIN_COVERAGE:.0%} → 기존 결과 유지, 실패 처리")
        return 1

    closes_all = {s: h["close"] for s, h in hist.items()}
    breadth = compute_breadth(closes_all)
    by_market = {m: compute_breadth({s: c for s, c in closes_all.items() if market_of.get(s) == m}) for m in ("KOSPI", "KOSDAQ")}
    sc = score_breadth(breadth)
    # 종목별 이동평균 거리 — 화면의 '20/50/200일선 아래 종목' 버튼용
    info_of = {yf_symbol(u["code"], u["market"]): u for u in universe}
    stocks = []
    for sym, closes in closes_all.items():
        u = info_of.get(sym)
        if not u:
            continue
        m = stock_metrics(closes)
        stocks.append({"code": u["code"], "name": u["name"], "market": u["market"], "mcap_b": u.get("mcap_b"),
                       "close": round(closes[-1], 2), **ma_distances(closes),
                       "high52": m.get("high52"), "low52": m.get("low52")})
    stocks.sort(key=lambda x: -(x["mcap_b"] or 0))
    print(f"  20/50/200일선 위: {breadth['pct_above_20']}% / {breadth['pct_above_50']}% / {breadth['pct_above_200']}%"
          f" · 신고가 {breadth['new_highs']} vs 신저가 {breadth['new_lows']} · 종합 {sc['score']}점 {sc['grade']} · 게이트 {'열림' if sc['gate_open'] else '차단'}")

    idx = {}
    proxies = fetch_price_history([p for pair in INDEX_PROXY.values() for p in pair], period="2y")
    for m, (isym, esym) in INDEX_PROXY.items():
        ih, eh = proxies.get(isym), proxies.get(esym)
        if ih:
            # 거래량은 ETF 대용 — 날짜를 맞춰 합친다
            vol_by_date = dict(zip(eh["dates"], eh["volume"])) if eh else {}
            vols = [vol_by_date.get(d, 0.0) for d in ih["dates"]]
            idx[m] = index_context(ih["close"], vols if eh else None, ih["dates"])
            idx[m]["volume_proxy"] = esym if eh else None
            print(f"  {m}: 고점대비 {idx[m].get('drawdown_from_high_pct')}% · 저점대비 +{idx[m].get('rebound_from_low_pct')}% ({idx[m].get('days_since_low')}일차)"
                  f" · FTD {idx[m].get('ftd') and idx[m]['ftd']['date'] or '없음'} · 분배일 {idx[m].get('distribution_days_25')}")

    as_of = max((h["dates"][-1] for h in hist.values()), default=now.strftime("%Y-%m-%d"))
    state = load_state(STATE_NAME, {"history": []})
    hist_pts = [p for p in (state.get("history") or []) if p.get("date") != as_of]
    hist_pts.append({"date": as_of, "score": sc["score"], "pct50": breadth["pct_above_50"], "pct200": breadth["pct_above_200"],
                     "highs": breadth["new_highs"], "lows": breadth["new_lows"]})
    hist_pts.sort(key=lambda p: p["date"])
    state["history"] = hist_pts[-MAX_HISTORY:]
    state["last_run"] = now.isoformat(timespec="minutes")
    save_state(STATE_NAME, state)

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "as_of": as_of,
        "universe": {"n": len(symbols), "fetched": len(hist), "basis": f"코스피 시총 상위 {TOP_KOSPI} + 코스닥 상위 {TOP_KOSDAQ} (ETF·ETN 제외)"},
        "breadth": breadth, "by_market": by_market,
        "score": sc["score"], "grade": sc["grade"], "hl_pct": sc.get("hl_pct"),
        "gate": {"open": sc["gate_open"], "reason": sc["reason"],
                 "rule": "종합 30점 미만 또는 50일선 위 25% 미만이면 차단"},
        "index": idx,
        "stocks": stocks,
        "history": [[p["date"], p["score"]] for p in state["history"] if p.get("score") is not None],
        "note": ("시장폭은 '종목을 고를 만한 장인지'를 보는 지표다. 게이트가 차단이면 아래 수급 스크리너 결과는 참고용으로만 본다. "
                 "FTD·분배일 거래량은 KODEX 200 / KOSDAQ 150 ETF 거래량을 대용으로 쓴다(지수 자체 거래량은 야후에 없음)."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"[OK] {OUTPUT_FILE}  기준일 {as_of} · 이력 {len(state['history'])}점")
    return 0


if __name__ == "__main__":
    sys.exit(main())
