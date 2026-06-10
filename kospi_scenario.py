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
    d = naver_kospi_daily(12)
    if d:
        k = sorted(d.keys())[-1]
        return k, d[k]
    return None, None


def naver_kospi_daily(days=90):
    """네이버 KOSPI 일별 종가 dict {YYYY-MM-DD: close} — 실증 회귀·보정용(신뢰 소스)."""
    import urllib.request
    import re
    from datetime import date, timedelta
    try:
        end = date.today().strftime("%Y%m%d")
        start = (date.today() - timedelta(days=days)).strftime("%Y%m%d")
        url = (f"https://api.finance.naver.com/siseJson.naver?symbol=KOSPI"
               f"&requestType=1&startTime={start}&endTime={end}&timeframe=day")
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0", "Referer": "https://finance.naver.com/"})
        txt = urllib.request.urlopen(req, timeout=10).read().decode("utf-8")
        rows = re.findall(r'\["(\d{8})",\s*[\d.]+,\s*[\d.]+,\s*[\d.]+,\s*([\d.]+)', txt)
        return {f"{d[:4]}-{d[4:6]}-{d[6:]}": float(c) for d, c in rows}
    except Exception as e:
        print(f"  [WARN] 네이버 KOSPI 일별 실패: {e}")
        return {}


def empirical_lag_betas(close_raw, np, window=20, backtest_n=12):
    """일별 실증 회귀: KOSPI(D) ~ SOX(P) + NDX(P) 무절편 (한국=미국 1일 시차 캐치업).
    현재 변동성 레짐을 반영해 베타를 재보정 → 진폭 과소예측 완화.
    반환: dict(sox, ndx, r2, resid_std, mae_bt, window, n) 또는 None."""
    kd_map = naver_kospi_daily(95)
    if len(kd_map) < window + 4:
        return None
    sret = close_raw["SOX"].dropna().pct_change()
    nret = close_raw["NDX"].dropna().pct_change()
    sr = {d.strftime("%Y-%m-%d"): float(v) for d, v in sret.items() if not np.isnan(v)}
    nr = {d.strftime("%Y-%m-%d"): float(v) for d, v in nret.items() if not np.isnan(v)}
    kdays = sorted(kd_map.keys())
    recs = []  # (y_kospi_D, s_prev, n_prev)
    for i in range(1, len(kdays)):
        D, P = kdays[i], kdays[i - 1]
        if P in sr and P in nr:
            recs.append((kd_map[D] / kd_map[P] - 1, sr[P], nr[P]))
    if len(recs) < window + 2:
        return None

    def fit(rows):
        Y = np.array([r[0] for r in rows])
        A = np.column_stack([[r[1] for r in rows], [r[2] for r in rows]])
        c, _, _, _ = np.linalg.lstsq(A, Y, rcond=None)
        return c, A, Y

    win = recs[-window:]
    coef, A, Y = fit(win)
    pred = A @ coef
    resid = Y - pred
    ss_res = float((resid ** 2).sum())
    ss_tot = float(((Y - Y.mean()) ** 2).sum()) or 1e-9
    r2 = 1 - ss_res / ss_tot
    resid_std = float(resid.std(ddof=1)) if len(resid) > 2 else float(resid.std())

    # 워크포워드 백테스트 MAE (직전 window일로 적합 → 다음날 예측)
    errs = []
    N = len(recs)
    for j in range(max(window, N - backtest_n), N):
        c2, _, _ = fit(recs[max(0, j - window):j])
        y = recs[j][0]
        p = c2[0] * recs[j][1] + c2[1] * recs[j][2]
        errs.append(abs((p - y) * 100))
    mae_bt = float(np.mean(errs)) if errs else None

    return {
        "sox": round(float(coef[0]), 3),
        "ndx": round(float(coef[1]), 3),
        "r2": round(r2, 3),
        "resid_std": round(resid_std, 4),
        "mae_bt": round(mae_bt, 2) if mae_bt is not None else None,
        "window": window,
        "n": len(recs),
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

    # 미국 지수 최근 1일 변화율 (KOSPI가 아직 반영 못한 야간 움직임 → 실시간 자동예측용)
    def _chg(col):
        s = close_raw[col].dropna()
        return round((float(s.iloc[-1]) / float(s.iloc[-2]) - 1) * 100, 2) if len(s) >= 2 else None
    change_pct = {"SOX": _chg("SOX"), "NDX": _chg("NDX"),
                  "SOXX": _chg("SOXX"), "QQQ": _chg("QQQ")}
    print(f"미국 1일 변화: SOX {change_pct['SOX']}% NDX {change_pct['NDX']}%")

    # 일별 실증 회귀 (현재 변동성 레짐 반영 → 진폭 과소예측 완화)
    emp = empirical_lag_betas(close_raw, np, window=20, backtest_n=12)
    if emp:
        print(f"실증 β(최근{emp['window']}일): SOX={emp['sox']} NDX={emp['ndx']} "
              f"R²={emp['r2']} 잔차σ={emp['resid_std']*100:.2f}% 백테스트MAE={emp['mae_bt']}%")
    else:
        print("  [WARN] 실증 회귀 실패 → 주간 stress β 폴백")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": asof,
        "current": now,
        "change_pct": change_pct,
        "beta": {
            "sox_weekly": round(bS, 3), "ndx_weekly": round(bN, 3),
            "sox_stress": round(bS_stress, 3), "ndx_stress": round(bN_stress, 3),
            "r_sox": round(rS, 2), "r_ndx": round(rN, 2),
        },
        "empirical": emp,  # 일별 실증 회귀(최근 레짐) — 프론트 중심예측·밴드의 1차 소스
        "default_targets": {"soxx": 460, "qqq": 650},
        "tail_multiplier": 1.8,
        "window_years": 3,
        "note": ("일별 실증 회귀(최근 20거래일, 무절편) β로 중심값 산출 — 현재 변동성 레짐 반영. "
                 "밴드=잔차 1σ, 꼬리=−1.5σ. 주간 stress β는 폴백."),
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
