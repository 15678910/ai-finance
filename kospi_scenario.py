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
US_CACHE = os.path.join(BASE_DIR, "docs", "us_index_cache.json")


def merge_us_cache(close_raw):
    """US 지수 일별종가를 캐시에 누적 병합 → yfinance 플립플롭(최신 바 유실) 방지.
    한 번 본 날짜는 영구 보존, 최신 날짜가 절대 stale로 되돌아가지 않음."""
    try:
        with open(US_CACHE, encoding="utf-8") as f:
            cache = json.load(f)
    except Exception:
        cache = {}
    for c in ("SOX", "NDX", "SOXX", "QQQ"):
        if c not in close_raw.columns:
            continue
        cache.setdefault(c, {})
        s = close_raw[c].dropna()
        for d, v in s.items():
            cache[c][d.strftime("%Y-%m-%d")] = round(float(v), 4)
    try:
        with open(US_CACHE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)
    except Exception as e:
        print(f"  [WARN] US 캐시 저장 실패: {e}")
    return cache


def cache_change(cache, col):
    """캐시의 최신 2거래일로 (변화율%, 최신날짜). yfinance보다 신뢰."""
    dd = sorted(cache.get(col, {}).keys())
    if len(dd) >= 2:
        latest, prev = dd[-1], dd[-2]
        p1, p0 = cache[col][latest], cache[col][prev]
        if p0:
            return round((p1 / p0 - 1) * 100, 2), latest
    return None, None


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


def futures_signal(np, pd, kospi_close_date, kospi_last_close, window=20):
    """선물 기반 장전 예측: 직전 KOSPI 마감 → 현재까지 나스닥선물(NQ=F) 야간 움직임.
    선물은 24시간 거래 → 미국 현물 마감 후 뉴스까지 반영 = 가장 신선한 선행신호.
    실증 R²=0.53(현물 0.28의 2배), 워크포워드 MAE 1.9%."""
    import yfinance as yf
    try:
        nq = yf.download("NQ=F", period="60d", interval="1h", progress=False)["Close"]
        if hasattr(nq, "columns"):  # CI yfinance는 단일종목도 DataFrame 반환 → Series로
            nq = nq.iloc[:, 0]
        nq = nq.dropna()
    except Exception as e:
        print(f"  [WARN] 선물 다운로드 실패: {e}")
        return None
    if len(nq) < 50:
        return None
    nq.index = nq.index.tz_convert("UTC") if nq.index.tz else nq.index.tz_localize("UTC")
    nq = nq.sort_index()

    def ab(ts):  # ts 이전(같거나) 마지막 선물값
        sub = nq[nq.index <= ts]
        return float(sub.iloc[-1]) if len(sub) else None

    kmap = naver_kospi_daily(80)
    kd = sorted(kmap.keys())
    recs = []  # (kospi_full_ret, fut_overnight)
    for i in range(1, len(kd)):
        D, P = kd[i], kd[i - 1]
        pc = kmap[P]
        f0 = ab(pd.Timestamp(f"{P} 06:30", tz="UTC"))  # 직전 KOSPI 마감(15:30 KST)
        f1 = ab(pd.Timestamp(f"{D} 00:00", tz="UTC"))   # KOSPI 개장(09:00 KST)
        if f0 and f1 and f0 > 0:
            r = f1 / f0 - 1
            if abs(r) < 0.2:
                recs.append((kmap[D] / pc - 1, r))
    if len(recs) < window + 2:
        return None

    win = recs[-window:]
    F = np.array([r[1] for r in win])
    Y = np.array([r[0] for r in win])
    denom = float(F @ F) or 1e-9
    beta = float((F @ Y) / denom)
    resid = Y - beta * F
    sigma = float(resid.std(ddof=1)) if len(resid) > 2 else float(resid.std())
    Fa = np.array([r[1] for r in recs])
    Ya = np.array([r[0] for r in recs])
    r2 = float(np.corrcoef(Ya, Fa)[0, 1] ** 2)

    errs = []
    N = len(recs)
    for j in range(max(window, N - 12), N):
        w = recs[j - window:j]
        Fw = np.array([x[1] for x in w])
        Yw = np.array([x[0] for x in w])
        bw = float((Fw @ Yw) / (float(Fw @ Fw) or 1e-9))
        errs.append(abs((bw * recs[j][1] - recs[j][0]) * 100))
    mae_bt = float(np.mean(errs)) if errs else None

    # 현재 라이브 신호: 직전 KOSPI 마감 이후 선물 움직임 → 다음 KOSPI 예측
    f_close = ab(pd.Timestamp(f"{kospi_close_date} 06:30", tz="UTC"))
    f_now = float(nq.iloc[-1])
    if not f_close or f_close <= 0:
        return None
    sig = f_now / f_close - 1
    pred = kospi_last_close * (1 + beta * sig)
    return {
        "fut_overnight_pct": round(sig * 100, 2),
        "beta": round(beta, 2),
        "predicted": round(pred, 0),
        "predicted_pct": round((pred / kospi_last_close - 1) * 100, 2),
        "sigma": round(sigma, 4),
        "r2": round(r2, 3),
        "mae_bt": round(mae_bt, 2) if mae_bt is not None else None,
        "window": window,
        "nq_now": round(f_now, 1),
        "nq_at_kospi_close": round(f_close, 1),
        "fut_asof_utc": nq.index[-1].strftime("%Y-%m-%d %H:%M UTC"),
        "kospi_base": round(kospi_last_close, 2),
    }


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


def nikkei_divergence(np, window=40):
    """한일 동행/괴리 진단(예측 아님). 같은 시간대(09:00 개장) → 당일 동시 비교.
    KOSPI 실측(같은 날 니케이) 베타로 기대치 산출 → 실제와의 괴리로 '한국 고유' 판별:
      괴리<0 = 한국이 일본보다 약함(외인 투매·환율 등 한국 악재)
      괴리>0 = 한국이 일본보다 강함(한국 호재)."""
    import yfinance as yf
    try:
        n = yf.download("^N225", period="4mo", interval="1d", progress=False, auto_adjust=True)["Close"]
        if hasattr(n, "columns"):  # DataFrame → Series
            n = n.iloc[:, 0]
        n = n.dropna()
    except Exception as e:
        print(f"  [WARN] 니케이 다운로드 실패: {e}")
        return None
    if len(n) < 15:
        return None
    kmap = naver_kospi_daily(80)
    nret = {d.strftime("%Y-%m-%d"): float(v) for d, v in n.pct_change().items() if not np.isnan(v)}
    nlvl = {d.strftime("%Y-%m-%d"): float(v) for d, v in n.items()}
    kd = sorted(kmap.keys())
    pairs = []  # (date, kospi_ret, nikkei_ret)
    for i in range(1, len(kd)):
        D, P = kd[i], kd[i - 1]
        if D in nret:
            pairs.append((D, kmap[D] / kmap[P] - 1, nret[D]))
    if len(pairs) < 12:
        return None
    win = pairs[-window:]
    K = np.array([p[1] for p in win])
    J = np.array([p[2] for p in win])
    denom = float(J @ J) or 1e-9
    beta_jp = float((J @ K) / denom)  # 무절편 동시간대 베타
    r = float(np.corrcoef(K, J)[0, 1])

    last = pairs[-1]
    ld, k_chg, j_chg = last
    expected = beta_jp * j_chg
    gap = k_chg - expected  # 한국 고유 성분
    if abs(gap) < 0.006:
        verdict, vkey = "한일 동행 (아시아 동반)", "sync"
    elif gap < 0:
        verdict, vkey = "한국 고유 약세 (외인·환율 등)", "kr_weak"
    else:
        verdict, vkey = "한국 고유 강세", "kr_strong"
    return {
        "date": ld,
        "kospi_chg": round(k_chg * 100, 2),
        "nikkei_chg": round(j_chg * 100, 2),
        "nikkei_level": round(nlvl.get(ld, 0), 2),
        "beta_jp": round(beta_jp, 2),
        "corr": round(r, 2),
        "expected_kospi_chg": round(expected * 100, 2),
        "gap_pct": round(gap * 100, 2),
        "verdict": verdict,
        "vkey": vkey,
        "n": len(pairs),
    }


def yen_watch(np):
    """USD/JPY(엔/달러) 감시 — 엔화는 아시아 위험의 선행지표. 급변 시 KOSPI 반전 신호(예측 아님·맥락)."""
    import yfinance as yf
    try:
        s = yf.download("JPY=X", period="12d", interval="1d", progress=False)["Close"]
        if hasattr(s, "columns"):  # 단일종목 DataFrame → Series
            s = s.iloc[:, 0]
        s = s.dropna()
    except Exception as e:
        print(f"  [WARN] 엔화 다운로드 실패: {e}")
        return None
    if len(s) < 3:
        return None
    cur = float(s.iloc[-1])
    chg1 = (cur / float(s.iloc[-2]) - 1) * 100
    chg5 = (cur / float(s.iloc[-6]) - 1) * 100 if len(s) >= 6 else None
    # 급변 경고: 1일 ±1% 또는 5일 ±2.5% (USD/JPY↑ = 엔 약세)
    alert = None
    if abs(chg1) >= 1.0 or (chg5 is not None and abs(chg5) >= 2.5):
        direction = "약세(엔↓)" if chg1 > 0 else "강세(엔↑)"
        alert = f"엔화 {direction} 급변 — 아시아 위험·BOJ 개입 주의"
    return {
        "usdjpy": round(cur, 2),
        "chg_1d": round(chg1, 2),
        "chg_5d": round(chg5, 2) if chg5 is not None else None,
        "alert": alert,
        "asof": s.index[-1].strftime("%Y-%m-%d"),
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

    # 미국 지수 최근 1일 변화율 — 캐시 사용(yfinance 플립플롭으로 최신 바 유실 방지)
    us_cache = merge_us_cache(close_raw)
    sox_chg, us_asof = cache_change(us_cache, "SOX")
    ndx_chg, _ = cache_change(us_cache, "NDX")
    soxx_chg, _ = cache_change(us_cache, "SOXX")
    qqq_chg, _ = cache_change(us_cache, "QQQ")
    # 미국 마감일이 KOSPI 기준일보다 같거나 이전이면 → 이미 KOSPI에 반영됨(다음 예측 신호 아님)
    absorbed = bool(us_asof and us_asof < asof)  # US 같은날=KOSPI 마감 후 거래 → 미반영
    change_pct = {"SOX": sox_chg, "NDX": ndx_chg, "SOXX": soxx_chg, "QQQ": qqq_chg,
                  "asof": us_asof, "kospi_asof": asof, "absorbed": absorbed}
    print(f"미국 최근({us_asof}): SOX {sox_chg}% NDX {ndx_chg}% "
          f"{'[KOSPI '+asof+' 이미 반영]' if absorbed else '[다음 거래일 미반영]'}")

    # 일별 실증 회귀 (현재 변동성 레짐 반영 → 진폭 과소예측 완화)
    emp = empirical_lag_betas(close_raw, np, window=20, backtest_n=12)
    if emp:
        print(f"실증 β(최근{emp['window']}일): SOX={emp['sox']} NDX={emp['ndx']} "
              f"R²={emp['r2']} 잔차σ={emp['resid_std']*100:.2f}% 백테스트MAE={emp['mae_bt']}%")
    else:
        print("  [WARN] 실증 회귀 실패 → 주간 stress β 폴백")

    # 선물 기반 장전 예측 (NQ=F 야간 → 다음 KOSPI) — 최고의 선행신호(R²0.53)
    futures = futures_signal(np, pd, asof, now["KOSPI"], window=20)
    if futures:
        print(f"선물 예측: NQ 야간 {futures['fut_overnight_pct']:+.2f}% → KOSPI {futures['predicted']:.0f} "
              f"({futures['predicted_pct']:+.2f}%) [β={futures['beta']} R²={futures['r2']} MAE={futures['mae_bt']}%]")
    else:
        print("  [WARN] 선물 신호 실패")

    # 한일 동행/괴리 진단 (예측 아님 — 원인 해석용)
    nikkei = nikkei_divergence(np, window=40)
    if nikkei:
        print(f"한일 진단({nikkei['date']}): KOSPI {nikkei['kospi_chg']:+.2f}% vs 니케이 {nikkei['nikkei_chg']:+.2f}% "
              f"(기대 {nikkei['expected_kospi_chg']:+.2f}%, 괴리 {nikkei['gap_pct']:+.2f}%) → {nikkei['verdict']}")

    # 엔/달러(USD/JPY) 감시 — 아시아 위험 선행지표 (BOJ·엔화 급변 맥락)
    yen = yen_watch(np)
    if yen:
        print(f"엔/달러: {yen['usdjpy']} ({yen['chg_1d']:+.2f}% 1d){' ⚠️ ' + yen['alert'] if yen['alert'] else ''}")

    # ※ 장중 보정(개장 시가갭→종가)은 '진짜 예측'이 아니라 이미 실현된 시가를 끌어쓰는
    #   정보 누출이라 제거됨. 예측은 개장 전(05:30 KST) 고정 → 마감 후 채점(prediction_tracker.py).

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
        "empirical": emp,  # 일별 실증 회귀(최근 레짐) — 슬라이더 가정도구·밴드
        "futures": futures,  # 선물 야간(NQ=F) 기반 장전 예측 — 최고의 선행신호(R²0.53)
        "nikkei_div": nikkei,  # 한일 동행/괴리 진단(예측 아님 — 원인 해석용)
        "yen": yen,  # 엔/달러 감시(아시아 위험 선행 — BOJ·엔화 급변 맥락)
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
