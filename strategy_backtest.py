"""
전략 백테스터 — 선물 신호 기반 매매 규칙의 과거 성과 검증
==========================================================
"예측이 맞다"와 "그 예측으로 돈을 번다"는 다르다. 실제 매매 규칙을
과거 데이터에 적용해 승률·누적수익·MDD를 측정 → 거래 가능성 검증.

규칙(KOSPI 종일, 전일종가→종가):
  A. 신호추종: 선물 야간 신호 방향대로 매수/매도
  B. 임계 매수: 선물 +0.5% 이상일 때만 매수(아니면 현금)
  C. Buy&Hold: 항상 매수 (벤치마크)

출력: docs/strategy_backtest.json
🚨 과거 성과가 미래를 보장하지 않음. 거래비용·슬리피지 미반영. 교육용.
"""

import json
import os
import re
import sys
import urllib.request
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "strategy_backtest.json")


def naver_kospi_daily(days=95):
    try:
        end = date.today().strftime("%Y%m%d")
        start = (date.today() - timedelta(days=days)).strftime("%Y%m%d")
        url = (f"https://api.finance.naver.com/siseJson.naver?symbol=KOSPI"
               f"&requestType=1&startTime={start}&endTime={end}&timeframe=day")
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0", "Referer": "https://finance.naver.com/"})
        txt = urllib.request.urlopen(req, timeout=12).read().decode("utf-8")
        rows = re.findall(r'\["(\d{8})",\s*[\d.]+,\s*[\d.]+,\s*[\d.]+,\s*([\d.]+)', txt)
        return {f"{d[:4]}-{d[4:6]}-{d[6:]}": float(c) for d, c in rows}
    except Exception as e:
        print(f"[ERROR] 네이버 KOSPI 실패: {e}")
        return {}


def metrics(rets, np):
    """일별 전략수익률 리스트 → 성과지표."""
    if not rets:
        return None
    r = np.array(rets)
    eq = np.cumprod(1 + r)
    cum = float(eq[-1] - 1) * 100
    peak = np.maximum.accumulate(eq)
    mdd = float(((eq - peak) / peak).min()) * 100
    traded = r[r != 0]
    win = float((traded > 0).mean()) * 100 if len(traded) else 0.0
    sharpe = float(r.mean() / r.std() * (252 ** 0.5)) if r.std() > 0 else 0.0
    return {
        "trades": int(len(traded)),
        "win_rate": round(win, 1),
        "cum_return": round(cum, 2),
        "avg_daily": round(float(traded.mean() * 100), 3) if len(traded) else 0.0,
        "mdd": round(mdd, 2),
        "sharpe": round(sharpe, 2),
    }


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
    print("  전략 백테스터 (선물 신호 매매 규칙)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    k = naver_kospi_daily(95)
    if len(k) < 20:
        print("[ERROR] KOSPI 데이터 부족")
        return 1

    nq = yf.download("NQ=F", period="60d", interval="1h", progress=False)["Close"]
    if hasattr(nq, "columns"):
        nq = nq.iloc[:, 0]
    nq = nq.dropna()
    nq.index = nq.index.tz_convert("UTC") if nq.index.tz else nq.index.tz_localize("UTC")
    nq = nq.sort_index()

    def ab(ts):
        sub = nq[nq.index <= ts]
        return float(sub.iloc[-1]) if len(sub) else None

    kd = sorted(k.keys())
    recs = []  # (date, kospi_ret, fut_signal)
    for i in range(1, len(kd)):
        D, P = kd[i], kd[i - 1]
        pc = k[P]
        f0 = ab(pd.Timestamp(f"{P} 06:30", tz="UTC"))
        f1 = ab(pd.Timestamp(f"{D} 00:00", tz="UTC"))
        if f0 and f1 and f0 > 0 and pc > 0:
            sig = f1 / f0 - 1
            if abs(sig) < 0.2:
                recs.append((D, k[D] / pc - 1, sig))
    if len(recs) < 10:
        print("[ERROR] 정렬 표본 부족")
        return 1

    y = [r[1] for r in recs]   # KOSPI 종일 수익률
    s = [r[2] for r in recs]   # 선물 야간 신호

    # A. 신호추종 (방향대로 매수/매도)
    a_ret = [yi if si > 0 else (-yi if si < 0 else 0) for yi, si in zip(y, s)]
    # B. 임계 매수 (+0.5% 이상만 매수)
    b_ret = [yi if si > 0.005 else 0 for yi, si in zip(y, s)]
    # C. Buy & Hold
    c_ret = list(y)

    strategies = [
        {"key": "follow", "name": "신호추종 (방향대로 매수/매도)", "m": metrics(a_ret, np)},
        {"key": "thresh", "name": "임계 매수 (선물 +0.5%↑만 매수)", "m": metrics(b_ret, np)},
        {"key": "hold", "name": "Buy & Hold (항상 매수·벤치마크)", "m": metrics(c_ret, np)},
    ]
    for st in strategies:
        mm = st["m"] or {}
        print(f"  {st['name']}: 거래 {mm.get('trades')} 승률 {mm.get('win_rate')}% "
              f"누적 {mm.get('cum_return')}% MDD {mm.get('mdd')}% Sharpe {mm.get('sharpe')}")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "period": f"{recs[0][0]} ~ {recs[-1][0]} ({len(recs)}거래일)",
        "n": len(recs),
        "strategies": strategies,
        "note": ("선물 야간 신호로 KOSPI 종일(전일종가→종가) 매매 가정. "
                 "거래비용·슬리피지·세금 미반영. 과거 성과는 미래 보장 아님. 교육용."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
