"""
KOSPI 시나리오 시뮬레이터 — 베타 회귀 기반 실시간 예측 엔진
==============================================================
SOX(필라델피아 반도체)·나스닥100 목표치 → KOSPI 예상 밴드.

- 3년 주간 수익률로 β(민감도) 회귀 + 하락장 조건부 stress β
- 현재가(KOSPI/SOX/SOXX/NDX/QQQ) 스냅샷
- 슬라이더 계산은 프론트엔드에서 즉시 수행 (이 파일은 β·현재가만 공급)

출력: docs/kospi_scenario.json
🚨 통계 추정. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "kospi_scenario.json")


def naver_kospi_latest():
    """네이버 일별시세 KOSPI 최신 (date, close) — yfinance ^KS11 stale 보정용."""
    import urllib.request
    import re
    from datetime import date, timedelta
    try:
        end = date.today().strftime("%Y%m%d")
        start = (date.today() - timedelta(days=12)).strftime("%Y%m%d")
        url = (f"https://api.finance.naver.com/siseJson.naver?symbol=KOSPI"
               f"&requestType=1&startTime={start}&endTime={end}&timeframe=day")
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0", "Referer": "https://finance.naver.com/"})
        txt = urllib.request.urlopen(req, timeout=10).read().decode("utf-8")
        rows = re.findall(r'\["(\d{8})",\s*[\d.]+,\s*[\d.]+,\s*[\d.]+,\s*([\d.]+)', txt)
        if rows:
            d, c = rows[-1]
            return f"{d[:4]}-{d[4:6]}-{d[6:]}", float(c)
    except Exception as e:
        print(f"  [WARN] 네이버 KOSPI 실패: {e}")
    return None, None


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
    print("  KOSPI 시나리오 시뮬레이터 — 베타 회귀")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    tk = ["^KS11", "^SOX", "^NDX", "SOXX", "QQQ"]
    raw = yf.download(tk, period="3y", interval="1d", progress=False, auto_adjust=True)
    px = raw["Close"].dropna().rename(columns={"^KS11": "KOSPI", "^SOX": "SOX", "^NDX": "NDX"})
    if len(px) < 60:
        print("[ERROR] 데이터 부족")
        return 1

    # 현재가는 각 지수의 '개별 최신 종가' (정렬 dropna는 미국 미마감 시 KOSPI 최신값을 버림)
    close_raw = raw["Close"].rename(columns={"^KS11": "KOSPI", "^SOX": "SOX", "^NDX": "NDX"})
    now = {c: round(float(close_raw[c].dropna().iloc[-1]), 2) for c in px.columns}
    kospi_asof = close_raw["KOSPI"].dropna().index[-1].strftime("%Y-%m-%d")
    # KOSPI yfinance stale 보정 (네이버 최신 종가)
    nk_date, nk_close = naver_kospi_latest()
    if nk_close and nk_date > kospi_asof:
        print(f"  [KOSPI 보정] yfinance {kospi_asof} → 네이버 {nk_date} {nk_close}")
        now["KOSPI"] = round(nk_close, 2)
        kospi_asof = nk_date
    asof = kospi_asof  # KOSPI 기준일 표시 (베타는 정렬 데이터로 계산)
    print(f"기준일(KOSPI) {asof} | 현재가: {now}")

    # 주간 리샘플 + 로그수익률
    wk = px.resample("W-FRI").last().dropna()
    wret = np.log(wk).diff().dropna()

    def beta(yv, xv):
        return float(np.polyfit(xv, yv, 1)[0])

    bS = beta(wret["KOSPI"], wret["SOX"])
    bN = beta(wret["KOSPI"], wret["NDX"])
    rS = float(np.corrcoef(wret["KOSPI"], wret["SOX"])[0, 1])
    rN = float(np.corrcoef(wret["KOSPI"], wret["NDX"])[0, 1])

    mS = wret["SOX"] < -0.02
    mN = wret["NDX"] < -0.02
    bS_stress = beta(wret["KOSPI"][mS], wret["SOX"][mS]) if mS.sum() >= 8 else bS
    bN_stress = beta(wret["KOSPI"][mN], wret["NDX"][mN]) if mN.sum() >= 8 else bN

    print(f"β_SOX={bS:.3f}(r={rS:.2f}) β_NDX={bN:.3f}(r={rN:.2f})")
    print(f"stress β_SOX={bS_stress:.3f} β_NDX={bN_stress:.3f}")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": asof,
        "current": now,
        "beta": {
            "sox_weekly": round(bS, 3), "ndx_weekly": round(bN, 3),
            "sox_stress": round(bS_stress, 3), "ndx_stress": round(bN_stress, 3),
            "r_sox": round(rS, 2), "r_ndx": round(rN, 2),
        },
        "default_targets": {"soxx": 460, "qqq": 650},
        "tail_multiplier": 1.8,
        "window_years": 3,
        "note": "주간 로그수익률 회귀 · 하락장(-2% 이하) 조건부 stress β · 패닉 시 β×1.8 오버슈팅 가정",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
