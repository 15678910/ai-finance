"""
KOSPI 예측 vs 실제 추적기
=========================
모델: KOSPI(D) ≈ KOSPI(전 거래일 P) × exp(0.5×(βs·ln(1+SOX_P) + βn·ln(1+NDX_P)))
  — 한국은 미국을 하루 늦게 따라감(P일 미국 마감 → D일 한국 반응).
매일 예측을 기록하고 실제 종가와 비교 → 적중도 산출.

데이터: 네이버 KOSPI 일별(신뢰) + yfinance SOX/NDX 일별, 베타는 kospi_scenario.json.
출력: docs/prediction_log.json
🚨 통계 추정. 투자 결정 단독 사용 금지.
"""

import json
import os
import re
import sys
import urllib.request
import math
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS, "prediction_log.json")


def naver_kospi_daily(days=30):
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


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    import numpy as np
    import yfinance as yf
    import warnings
    warnings.filterwarnings("ignore")

    print("=" * 55)
    print("  KOSPI 예측 vs 실제 추적기")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    # 베타 (대시보드와 동일 — 0.5×stress 블렌드)
    try:
        with open(os.path.join(DOCS, "kospi_scenario.json"), encoding="utf-8") as f:
            beta = json.load(f)["beta"]
        bSs, bNs = beta["sox_stress"], beta["ndx_stress"]
    except Exception:
        bSs, bNs = 0.337, 0.502
    print(f"베타: SOX stress {bSs} · NDX stress {bNs}")

    kospi = naver_kospi_daily(30)
    if len(kospi) < 3:
        print("[ERROR] KOSPI 데이터 부족")
        return 1

    # 미국 SOX/NDX 일별 수익률
    raw = yf.download(["^SOX", "^NDX"], period="2mo", interval="1d", progress=False, auto_adjust=True)
    cl = raw["Close"].dropna()
    sox_ret = cl["^SOX"].pct_change()
    ndx_ret = cl["^NDX"].pct_change()
    sr = {d.strftime("%Y-%m-%d"): float(v) for d, v in sox_ret.items() if not np.isnan(v)}
    nr = {d.strftime("%Y-%m-%d"): float(v) for d, v in ndx_ret.items() if not np.isnan(v)}

    kdays = sorted(kospi.keys())
    entries = []
    for i in range(1, len(kdays)):
        D, P = kdays[i], kdays[i - 1]
        if P not in sr or P not in nr:
            continue  # 미국 휴장 등
        blend = 0.5 * (bSs * math.log(1 + sr[P]) + bNs * math.log(1 + nr[P]))
        pred = kospi[P] * math.exp(blend)
        actual = kospi[D]
        err = (actual / pred - 1) * 100
        pred_dir = pred >= kospi[P]
        act_dir = actual >= kospi[P]
        entries.append({
            "date": D,
            "base": round(kospi[P], 2),
            "sox_ret": round(sr[P] * 100, 2),
            "ndx_ret": round(nr[P] * 100, 2),
            "predicted": round(pred, 0),
            "pred_pct": round((pred / kospi[P] - 1) * 100, 2),
            "actual": round(actual, 2),
            "actual_pct": round((actual / kospi[P] - 1) * 100, 2),
            "error_pct": round(err, 2),
            "abs_error": round(abs(err), 2),
            "dir_correct": (pred_dir == act_dir),
        })

    recent = entries[-12:]
    n = len(recent)
    mae = round(sum(e["abs_error"] for e in recent) / n, 2) if n else None
    dir_hit = round(sum(1 for e in recent if e["dir_correct"]) / n * 100) if n else None

    for e in recent[-6:]:
        print(f"  {e['date']} 예측 {e['predicted']:.0f}({e['pred_pct']:+.1f}%) vs 실제 {e['actual']:.0f}({e['actual_pct']:+.1f}%) 오차 {e['error_pct']:+.2f}%")
    print(f"\n최근 {n}일 MAE {mae}% · 방향 적중 {dir_hit}%")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "model": "KOSPI(D) = 전일종가 × exp(0.5×(βs·SOX + βn·NDX)) — 미국 1일 시차 캐치업",
        "beta": {"sox_stress": bSs, "ndx_stress": bNs},
        "entries": recent,
        "accuracy": {"mae_pct": mae, "dir_hit_rate": dir_hit, "n": n},
        "note": "예측=전일 KOSPI 종가에 미국 야간 움직임을 베타 환산. 실제=당일 종가. 통계 추정.",
    }
    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} ({n}건)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
