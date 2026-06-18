"""
추세·모멘텀 오버레이 백테스트 — 선물 단일신호의 하방편향을 보정할 수 있나?
==============================================================================
문제: 6/16~18 SK·KOSPI가 3일 연속 '하락 예측 → 외국인 매수로 급등'.
가설: 선물(미국)만으론 한국 추세장을 못 잡는다. 최근 가격 모멘텀(전 거래일까지의
      N일 추세, 누설 없음)을 2번째 신호로 넣으면 하방편향이 줄어들 것이다.

방법(워크포워드, 해당일 제외):
  · 기준(A): 종가수익률 ~ 선물야간(fov)         [현재 운영 모델]
  · 오버레이(B): 종가수익률 ~ 선물야간 + 모멘텀(mom3)  [2-요인]
  → 종가 MAE·방향적중을 비교. B가 분명히 나을 때만 채택(과적합 경계).

대상: SK하이닉스·삼성전자·KOSPI. 출력: docs/momentum_backtest.json
🚨 채택 여부는 결과로 판단. 좋아지지 않으면 기존 모델 유지.
"""

import json
import os
import re
import sys
import urllib.request
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "momentum_backtest.json")
W = 25

TARGETS = [("SK하이닉스", "000660"), ("삼성전자", "005930"), ("KOSPI", "KOSPI")]


def naver_ohlc(code, days=200):
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
        print(f"[ERROR] 네이버 OHLC 실패({code}): {e}")
        return []


def build_recs(bars, ab, pd):
    """선물야간(fov) + 모멘텀(mom3, 전 거래일까지 3일 추세, 누설 없음) 정렬 표본."""
    recs = []
    for i in range(4, len(bars)):
        D, P = bars[i][0], bars[i - 1][0]
        pc = bars[i - 1][4]
        c = bars[i][4]
        f0 = ab(pd.Timestamp(f"{P} 06:30", tz="UTC"))
        f1 = ab(pd.Timestamp(f"{D} 00:00", tz="UTC") - pd.Timedelta(hours=2.0))
        if not (f0 and f1 and f0 > 0 and pc > 0):
            continue
        fo = f1 / f0 - 1
        if abs(fo) >= 0.2:
            continue
        mom3 = bars[i - 1][4] / bars[i - 4][4] - 1  # 기준일까지의 3거래일 추세 (개장 전 이미 알려진 값)
        recs.append({"D": D, "base": pc, "c_ret": c / pc - 1, "fov": fo, "mom3": mom3})
    return recs


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
    print("=" * 60)
    print("  추세·모멘텀 오버레이 백테스트 (선물 vs 선물+모멘텀)")
    print("=" * 60)

    nq = yf.download("NQ=F", period="180d", interval="1h", progress=False)["Close"]
    if hasattr(nq, "columns"):
        nq = nq.iloc[:, 0]
    nq = nq.dropna()
    nq.index = nq.index.tz_convert("UTC") if nq.index.tz else nq.index.tz_localize("UTC")
    nq = nq.sort_index()

    def ab(ts):
        sub = nq[nq.index <= ts]
        return float(sub.iloc[-1]) if len(sub) else None

    results = []
    for name, code in TARGETS:
        bars = naver_ohlc(code, 220)
        if len(bars) < W + 30:
            print(f"  [SKIP] {name}: 데이터 부족")
            continue
        bars.sort(key=lambda x: x[0])
        recs = build_recs(bars, ab, pd)
        if len(recs) < W + 10:
            print(f"  [SKIP] {name}: 정렬표본 부족({len(recs)})")
            continue

        a_err, b_err, a_dir, b_dir, n = [], [], 0, 0, 0
        for j in range(W, len(recs)):
            win = recs[j - W:j]
            r = recs[j]
            F = np.array([x["fov"] for x in win])
            M = np.array([x["mom3"] for x in win])
            Y = np.array([x["c_ret"] for x in win])
            # A: 절편 없는 단일 회귀 (운영 모델과 동일)
            bC = float((F @ Y) / (float(F @ F) or 1e-9))
            pa = bC * r["fov"]
            # B: 2-요인 최소제곱 (fov, mom3)
            X = np.column_stack([F, M])
            try:
                coef, *_ = np.linalg.lstsq(X, Y, rcond=None)
            except Exception:
                coef = np.array([bC, 0.0])
            pb = float(coef[0] * r["fov"] + coef[1] * r["mom3"])
            act = r["c_ret"]
            a_err.append(abs(pa - act))
            b_err.append(abs(pb - act))
            a_dir += int((pa >= 0) == (act >= 0))
            b_dir += int((pb >= 0) == (act >= 0))
            n += 1

        mae_a = round(float(np.mean(a_err)) * 100, 3)
        mae_b = round(float(np.mean(b_err)) * 100, 3)
        dir_a = round(a_dir / n * 100, 1)
        dir_b = round(b_dir / n * 100, 1)
        better = (mae_b < mae_a - 0.02) and (dir_b >= dir_a)  # MAE 개선 + 방향 비악화
        verdict = "✅ 오버레이 우위" if better else ("➖ 유의차 없음" if abs(mae_b - mae_a) < 0.05 else "❌ 오버레이 악화")
        results.append({"name": name, "n": n, "mae_a": mae_a, "mae_b": mae_b,
                        "dir_a": dir_a, "dir_b": dir_b, "verdict": verdict})
        print(f"  {name:10s} n={n:3d} | MAE A {mae_a:.3f}% → B {mae_b:.3f}% | "
              f"방향 A {dir_a:.0f}% → B {dir_b:.0f}% | {verdict}")

    win_cnt = sum(1 for r in results if r["verdict"].startswith("✅"))
    overall = ("채택 권고 — 다수 대상에서 모멘텀이 정확도 개선"
               if win_cnt >= 2 else
               "기존 모델 유지 권고 — 모멘텀 추가가 일관된 개선을 못 줌(과적합 위험). "
               "대신 아시아 반도체 동조·디커플링 경고로 신뢰도만 표시.")
    print(f"\n  종합: {overall}")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "window": W, "results": results, "win_count": win_cnt,
        "recommendation": overall,
        "note": ("A=선물야간 단일(운영 모델), B=선물+3일모멘텀 2요인. 워크포워드(해당일 제외). "
                 "모멘텀은 기준일까지의 과거 종가만 사용(누설 없음). "
                 "B 채택은 2개 이상 대상에서 MAE 개선+방향 비악화일 때만."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
