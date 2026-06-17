"""
SK하이닉스 OHLC 예측기 — 개장 전 고정 + 마감 후 채점 (누적 성적표)
====================================================================
KOSPI 시나리오와 동일 철학: '진짜 예측'만 평가한다.
  · 개장 전(07:30 KST, 미국 마감 직후) 나스닥선물 야간신호로 다음 거래일 OHLC 예측
  · 한 번 고정하면 장중 시장 따라 수정 금지(freeze) — today 블록 persist
  · 마감 후 실제 종가로 워크포워드 채점(해당일 제외) → 종가 MAE·방향·범위 적중 누적

엔진: 나스닥선물(NQ=F) 24H 야간(직전 SK 마감 06:30 UTC → 07:30 KST) × 워크포워드 β.
참고: 마이크론(MU, 메모리 경쟁사)·브로드컴·SOXX·엔비디아 동향.

출력: docs/sk_hynix_forecast.json
🚨 통계 추정. 투자 결정 단독 사용 금지.
"""

import json
import os
import re
import sys
import urllib.request
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "sk_hynix_forecast.json")
CODE = "000660"  # SK하이닉스
W = 25           # 워크포워드 윈도우


def naver_ohlc(code, days=120):
    """네이버 일별 OHLC → [(date, o, h, l, c)] 오름차순."""
    try:
        end = date.today().strftime("%Y%m%d")
        start = (date.today() - timedelta(days=days)).strftime("%Y%m%d")
        url = (f"https://api.finance.naver.com/siseJson.naver?symbol={code}"
               f"&requestType=1&startTime={start}&endTime={end}&timeframe=day")
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0", "Referer": "https://finance.naver.com/"})
        txt = urllib.request.urlopen(req, timeout=12).read().decode("utf-8")
        rows = re.findall(r'\["(\d{8})",\s*([\d.]+),\s*([\d.]+),\s*([\d.]+),\s*([\d.]+)', txt)
        return [(f"{d[:4]}-{d[4:6]}-{d[6:]}", float(o), float(h), float(l), float(c))
                for d, o, h, l, c in rows]
    except Exception as e:
        print(f"[ERROR] 네이버 OHLC 실패: {e}")
        return []


def _round_tick(p):
    """KRX 호가단위: 50만원 이상 1000원."""
    return round(p / 1000) * 1000


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

    now = datetime.now(KST)
    print("=" * 55)
    print("  SK하이닉스 OHLC 예측 — 개장전 고정 + 마감후 채점")
    print(f"  KST: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    bars = naver_ohlc(CODE, 120)
    if len(bars) < W + 5:
        print("[ERROR] SK하이닉스 데이터 부족")
        return 1
    bars.sort(key=lambda x: x[0])
    # 장중 미확정 바 제외 — 오늘은 마감(15:40 KST) 후에만 채점(실제 종가 오류 방지)
    if (now.hour * 100 + now.minute) < 1540 and bars and bars[-1][0] == now.strftime("%Y-%m-%d"):
        bars = bars[:-1]
    last_date, lo_o, lo_h, lo_l, last_close = bars[-1]
    print(f"SK하이닉스 최근 {last_date}: 시{lo_o:.0f} 고{lo_h:.0f} 저{lo_l:.0f} 종{last_close:.0f}")

    # 나스닥선물 1시간봉 (24H 야간 신호)
    nq = yf.download("NQ=F", period="60d", interval="1h", progress=False)["Close"]
    if hasattr(nq, "columns"):       # CI yfinance 단일종목도 DataFrame → Series
        nq = nq.iloc[:, 0]
    nq = nq.dropna()
    nq.index = nq.index.tz_convert("UTC") if nq.index.tz else nq.index.tz_localize("UTC")
    nq = nq.sort_index()

    def ab(ts):  # ts 이전(같거나) 마지막 선물값
        sub = nq[nq.index <= ts]
        return float(sub.iloc[-1]) if len(sub) else None

    # 역사 표본: 직전마감(06:30 UTC) → 07:30 KST(=D 00:00 UTC-1.5h) 선물 야간 + 실제 OHLC
    recs = []
    for i in range(1, len(bars)):
        D, P = bars[i][0], bars[i - 1][0]
        pc = bars[i - 1][4]
        o, h, l, c = bars[i][1], bars[i][2], bars[i][3], bars[i][4]
        f0 = ab(pd.Timestamp(f"{P} 06:30", tz="UTC"))
        f1 = ab(pd.Timestamp(f"{D} 00:00", tz="UTC") - pd.Timedelta(hours=1.5))  # 07:30 KST D
        if f0 and f1 and f0 > 0 and pc > 0:
            fo = f1 / f0 - 1
            if abs(fo) < 0.2:
                bt, bb = max(o, c), min(o, c)
                recs.append({"D": D, "base": pc, "o": o, "h": h, "l": l, "c": c,
                             "c_ret": c / pc - 1, "g_ret": o / pc - 1, "fov": fo,
                             "uw": h / bt - 1, "dw": l / bb - 1})
    if len(recs) < W + 2:
        print("[ERROR] 선물 정렬 표본 부족")
        return 1

    def betas(win):
        F = np.array([r["fov"] for r in win])
        den = float(F @ F) or 1e-9
        bC = float((F @ np.array([r["c_ret"] for r in win])) / den)
        bG = float((F @ np.array([r["g_ret"] for r in win])) / den)
        uw = float(np.median([r["uw"] for r in win]))
        dw = float(np.median([r["dw"] for r in win]))
        return bC, bG, uw, dw

    # 최신 윈도우 베타 (오늘 예측·표시용)
    b_close, b_gap, up_wick, dn_wick = betas(recs[-W:])
    G2 = np.array([r["g_ret"] for r in recs[-W:]])
    Yc2 = np.array([r["c_ret"] for r in recs[-W:]])
    b_oc = float((G2 @ Yc2) / (float(G2 @ G2) or 1e-9))
    sigma_c = float((Yc2 - b_close * np.array([r["fov"] for r in recs[-W:]])).std(ddof=1))
    Fa, Ya = np.array([r["fov"] for r in recs]), np.array([r["c_ret"] for r in recs])
    r2 = float(np.corrcoef(Ya, Fa)[0, 1] ** 2)

    # ── 워크포워드 채점 (해당일 제외 직전 W일로 적합 → 종가 예측 vs 실제) ──
    entries = []
    for j in range(W, len(recs)):
        r = recs[j]
        bC, bG, uw, dw = betas(recs[j - W:j])
        base, fov = r["base"], r["fov"]
        p_open = base * (1 + bG * fov)
        p_close = base * (1 + bC * fov)
        p_high = max(p_open, p_close) * (1 + uw)
        p_low = min(p_open, p_close) * (1 + dw)
        actual = r["c"]
        err = (actual / p_close - 1) * 100
        entries.append({
            "date": r["D"], "base": round(base, 0),
            "fut_overnight": round(fov * 100, 2),
            "pred_open": _round_tick(p_open), "pred_close": _round_tick(p_close),
            "pred_high": _round_tick(p_high), "pred_low": _round_tick(p_low),
            "actual_open": round(r["o"], 0), "actual_high": round(r["h"], 0),
            "actual_low": round(r["l"], 0), "actual_close": round(actual, 0),
            "actual_close_pct": round(r["c_ret"] * 100, 2),
            "close_err": round(err, 2), "abs_err": round(abs(err), 2),
            "dir_ok": bool((p_close >= base) == (actual >= base)),
            "range_hit": bool(p_low <= actual <= p_high),
        })
    entries = entries[-15:]
    n = len(entries)
    mae = round(sum(e["abs_err"] for e in entries) / n, 2) if n else None
    dir_hit = round(sum(1 for e in entries if e["dir_ok"]) / n * 100) if n else None
    range_hit = round(sum(1 for e in entries if e["range_hit"]) / n * 100) if n else None
    errs = [e["abs_err"] for e in entries]
    mae_bt = round(sum(errs) / len(errs), 2) if errs else None

    # ── 다음 세션 예측 (개장 전 고정 / 잠정) ──
    def compute_today(status, target_label):
        f0 = ab(pd.Timestamp(f"{last_date} 06:30", tz="UTC"))
        f_now = float(nq.iloc[-1])
        sig = (f_now / f0 - 1) if (f0 and f0 > 0) else 0.0
        op = last_close * (1 + b_gap * sig)
        cl = last_close * (1 + b_close * sig)
        hi = max(op, cl) * (1 + up_wick)
        lo = min(op, cl) * (1 + dn_wick)
        return {
            "target_date": target_label, "status": status,
            "locked_at": now.strftime("%Y-%m-%d %H:%M KST"),
            "base_date": last_date, "base_close": round(last_close, 0),
            "fut_overnight_pct": round(sig * 100, 2),
            "fut_asof_utc": nq.index[-1].strftime("%Y-%m-%d %H:%M UTC"),
            "open": _round_tick(op), "high": _round_tick(hi),
            "low": _round_tick(lo), "close": _round_tick(cl),
            "open_pct": round((op / last_close - 1) * 100, 2),
            "close_pct": round((cl / last_close - 1) * 100, 2),
        }

    # 직전 출력 로드 (freeze 유지용)
    prev_today = {}
    try:
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            prev_today = (json.load(f) or {}).get("today") or {}
    except Exception:
        prev_today = {}

    today_str = now.strftime("%Y-%m-%d")
    hm = now.hour * 100 + now.minute
    LOCK_LO, LOCK_HI, CLOSE_HM = 700, 900, 1530  # 07:00~09:00 개장전 윈도우(07:30 신호), 15:30 마감

    if hm < CLOSE_HM:
        # 오늘 세션 대상 — 개장 전 고정 또는 장중 고정 유지(수정 금지)
        if prev_today.get("status") == "locked" and prev_today.get("target_date") == today_str:
            today = prev_today
            print(f"  오늘({today_str}) 예측 고정 유지 — 장중 수정 안 함")
        elif LOCK_LO <= hm < LOCK_HI:
            today = compute_today("locked", today_str)
            print(f"  ▶ {today_str} 개장 전 고정: 종가 {today['close']:.0f} ({today['close_pct']:+.2f}%)")
        else:
            today = compute_today("provisional", today_str)
            print(f"  ▶ {today_str} 잠정(07:30 확정 전): 종가 {today['close']:.0f}")
    else:
        # 마감 후 ~ 익일 새벽: 다음 거래일 잠정(내일 07:30 확정 예정)
        today = compute_today("provisional", "다음 거래일")
        print(f"  ▶ 다음 거래일 잠정(내일 07:30 확정 예정): 종가 {today['close']:.0f}")

    # 이벤트 영향권 경고(FOMC/BOJ ±1거래일) — 점 예측은 못 고치므로 신뢰도만 명시
    try:
        from event_risk import event_risk_block
        _ev_date = today_str if today.get("target_date") in (None, "다음 거래일") else today["target_date"]
        today = dict(today)
        today["event_risk"] = event_risk_block(_ev_date, window_days=1)
        if today["event_risk"]["flag"]:
            print(f"  ⚠️ 이벤트 영향권: {[e['title'] for e in today['event_risk']['events']]}")
    except Exception as e:
        print(f"  [WARN] event_risk 실패: {e}")

    # 미국 반도체 동료 맥락
    peers = {}
    try:
        praw = yf.download(["MU", "AVGO", "SOXX", "NVDA"], period="6d", interval="1d",
                           progress=False, auto_adjust=True)["Close"].dropna()
        for t, nm in {"MU": "마이크론", "AVGO": "브로드컴", "SOXX": "SOXX반도체", "NVDA": "엔비디아"}.items():
            s = praw[t].dropna()
            if len(s) >= 2:
                peers[nm] = round((float(s.iloc[-1]) / float(s.iloc[-2]) - 1) * 100, 2)
    except Exception as e:
        print(f"  [WARN] 동료 수집 실패: {e}")

    print(f"  채점({n}일): 종가 MAE ±{mae}% · 방향 {dir_hit}% · 범위 적중 {range_hit}%")

    output = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "name": "SK하이닉스", "code": CODE,
        "today": today,
        "entries": entries,
        "accuracy": {"close_mae_pct": mae, "dir_hit_rate": dir_hit,
                     "range_hit_rate": range_hit, "n": n},
        "model": {
            "beta_close": round(b_close, 2), "beta_gap": round(b_gap, 2),
            "beta_open_to_close": round(b_oc, 2), "sigma_close": round(sigma_c, 4),
            "up_wick": round(up_wick, 4), "dn_wick": round(dn_wick, 4),
            "r2": round(r2, 3), "mae_bt": mae_bt, "window": W,
        },
        "peers": peers,
        "last_ohlc": {"open": lo_o, "high": lo_h, "low": lo_l, "close": last_close},
        "note": ("개장 전(07:30 KST) 나스닥선물 야간신호로 다음 거래일 OHLC를 고정 → 장중 수정 없음 → "
                 "마감 후 실제 종가로 워크포워드 채점. 범위 적중=실제 종가가 예측 고저 안에 들어온 비율. "
                 "마이크론(메모리 경쟁사) 동향이 가장 밀접. 통계 추정."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
