"""
SK하이닉스 OHLC 예측기 (시가·고가·저가·종가)
================================================
KOSPI 방법론을 개별 종목에 적용 — 미국 마감 → 다음날 한국 반응(1일 시차).
엔진: 나스닥 선물(NQ=F) 24H 야간 움직임 = 가장 신선한 선행신호.
참고: 마이크론(MU, 메모리 경쟁사)·브로드컴(AVGO)·SOXX 반도체 동향.

실증(워크포워드):
  - 선물야간 ↔ SK 시가갭 corr 0.79 / 종일 corr 0.67(R²0.45, β~3.4)
  - 시가갭 ↔ 종일 corr 0.80 → 개장 후 종가오차 4.7%→2.7%
  - 일중 범위(고저폭 평균 5.3%)로 고가·저가 추정

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

    print("=" * 55)
    print("  SK하이닉스 OHLC 예측기")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    bars = naver_ohlc(CODE, 120)
    if len(bars) < 30:
        print("[ERROR] SK하이닉스 데이터 부족")
        return 1
    bars.sort(key=lambda x: x[0])
    last = bars[-1]
    last_date, lo_o, lo_h, lo_l, last_close = last
    print(f"SK하이닉스 최근 {last_date}: 시{lo_o:.0f} 고{lo_h:.0f} 저{lo_l:.0f} 종{last_close:.0f}")

    # 나스닥 선물 1시간봉 (24H 야간 신호)
    nq = yf.download("NQ=F", period="60d", interval="1h", progress=False)["Close"]
    if hasattr(nq, "columns"):  # CI yfinance는 단일종목도 DataFrame 반환 → Series로
        nq = nq.iloc[:, 0]
    nq = nq.dropna()
    nq.index = nq.index.tz_convert("UTC") if nq.index.tz else nq.index.tz_localize("UTC")
    nq = nq.sort_index()

    def ab(ts):
        sub = nq[nq.index <= ts]
        return float(sub.iloc[-1]) if len(sub) else None

    # 역사 표본: (full_ret, gap_ret, fut_overnight, up_wick, dn_wick)
    recs = []
    W = 25
    for i in range(1, len(bars)):
        D, P = bars[i][0], bars[i - 1][0]
        pc = bars[i - 1][4]
        o, h, l, c = bars[i][1], bars[i][2], bars[i][3], bars[i][4]
        f0 = ab(pd.Timestamp(f"{P} 06:30", tz="UTC"))
        f1 = ab(pd.Timestamp(f"{D} 00:00", tz="UTC"))
        if f0 and f1 and f0 > 0 and pc > 0:
            fo = f1 / f0 - 1
            if abs(fo) < 0.2:
                body_top, body_bot = max(o, c), min(o, c)
                recs.append((c / pc - 1, o / pc - 1, fo,
                             h / body_top - 1, l / body_bot - 1))
    if len(recs) < W + 2:
        print("[ERROR] 선물 정렬 표본 부족")
        return 1

    win = recs[-W:]
    F = np.array([r[2] for r in win])
    Yc = np.array([r[0] for r in win])  # 종일
    Yg = np.array([r[1] for r in win])  # 시가갭
    denom = float(F @ F) or 1e-9
    b_close = float((F @ Yc) / denom)
    b_gap = float((F @ Yg) / denom)
    resid_c = Yc - b_close * F
    sigma_c = float(resid_c.std(ddof=1))
    # 윅(고가/저가) 중앙값
    up_wick = float(np.median([r[3] for r in win]))
    dn_wick = float(np.median([r[4] for r in win]))
    # 시가갭→종가 베타(개장 후 보정용)
    G2 = np.array([r[1] for r in win])
    b_oc = float((G2 @ Yc) / (float(G2 @ G2) or 1e-9))

    # 전체 R² + 워크포워드 MAE(종가)
    Fa = np.array([r[2] for r in recs])
    Ya = np.array([r[0] for r in recs])
    r2 = float(np.corrcoef(Ya, Fa)[0, 1] ** 2)
    errs = []
    N = len(recs)
    for j in range(max(W, N - 12), N):
        w = recs[j - W:j]
        Fw = np.array([x[2] for x in w])
        Yw = np.array([x[0] for x in w])
        bw = float((Fw @ Yw) / (float(Fw @ Fw) or 1e-9))
        errs.append(abs((bw * recs[j][2] - recs[j][0]) * 100))
    mae_bt = float(np.mean(errs)) if errs else None

    # 현재 라이브 선물신호 (직전 SK 마감 06:30 UTC → 현재)
    f_close = ab(pd.Timestamp(f"{last_date} 06:30", tz="UTC"))
    f_now = float(nq.iloc[-1])
    sig = (f_now / f_close - 1) if (f_close and f_close > 0) else 0.0

    # 예측 OHLC (다음 거래일)
    open_pred = last_close * (1 + b_gap * sig)
    close_pred = last_close * (1 + b_close * sig)
    high_pred = max(open_pred, close_pred) * (1 + up_wick)
    low_pred = min(open_pred, close_pred) * (1 + dn_wick)

    # 미국 반도체 동료 맥락 (마이크론·브로드컴·SOXX 최근 등락)
    peers = {}
    try:
        praw = yf.download(["MU", "AVGO", "SOXX", "NVDA"], period="6d", interval="1d",
                           progress=False, auto_adjust=True)["Close"].dropna()
        pmap = {"MU": "마이크론", "AVGO": "브로드컴", "SOXX": "SOXX반도체", "NVDA": "엔비디아"}
        for t, nm in pmap.items():
            s = praw[t].dropna()
            if len(s) >= 2:
                peers[nm] = round((float(s.iloc[-1]) / float(s.iloc[-2]) - 1) * 100, 2)
    except Exception as e:
        print(f"  [WARN] 동료 수집 실패: {e}")

    print(f"선물 야간 {sig*100:+.2f}% (β종가 {b_close:.2f} β시가 {b_gap:.2f}) R²{r2:.2f} MAE±{mae_bt}%")
    print(f"예측: 시{_round_tick(open_pred):.0f} 고{_round_tick(high_pred):.0f} "
          f"저{_round_tick(low_pred):.0f} 종{_round_tick(close_pred):.0f}")
    print(f"동료: {peers}")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "name": "SK하이닉스", "code": CODE,
        "base_date": last_date,
        "base_close": round(last_close, 0),
        "fut_overnight_pct": round(sig * 100, 2),
        "fut_asof_utc": nq.index[-1].strftime("%Y-%m-%d %H:%M UTC"),
        "forecast": {
            "open": _round_tick(open_pred),
            "high": _round_tick(high_pred),
            "low": _round_tick(low_pred),
            "close": _round_tick(close_pred),
            "open_pct": round((open_pred / last_close - 1) * 100, 2),
            "close_pct": round((close_pred / last_close - 1) * 100, 2),
        },
        "model": {
            "beta_close": round(b_close, 2), "beta_gap": round(b_gap, 2),
            "beta_open_to_close": round(b_oc, 2),
            "sigma_close": round(sigma_c, 4),
            "up_wick": round(up_wick, 4), "dn_wick": round(dn_wick, 4),
            "r2": round(r2, 3), "mae_bt": round(mae_bt, 2) if mae_bt is not None else None,
            "window": W,
        },
        "peers": peers,
        "last_ohlc": {"open": lo_o, "high": lo_h, "low": lo_l, "close": last_close},
        "note": ("나스닥 선물 24H 야간신호 → 다음 거래일 시가·종가 예측, 일중 범위로 고가·저가 추정. "
                 "마이크론(메모리 경쟁사) 동향이 SK하이닉스와 가장 밀접. 통계 추정."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
