"""
KOSPI 개장 전 예측 vs 실제 추적기 (듀얼모델)
=============================================
'진짜 예측'만 채점한다 — 장 시작 전(07:00 KST, 미국 마감 직후)에 확정 가능한 정보만으로
당일 KOSPI 종가를 예측하고, 마감 후 실제 종가로 채점. 장중 보정(시가갭) 같은 정보 누출 금지.

두 모델을 매일 워크포워드(해당일 제외, 직전 N일로만 적합)로 예측·채점:
  · 회귀  : KOSPI(D) = 전일종가 × (1 + βs·SOX(P) + βn·NDX(P))  — 미국 1일 시차 캐치업
  · 선물  : KOSPI(D) = 전일종가 × (1 + βf·NQ야간)             — 나스닥선물 24H(직전 마감→07:00)
  · 앙상블: 두 모델 평균
누적 채점으로 모델별 MAE·방향 적중률을 산출 → 어느 쪽이 우수한지 데이터로 판정(연구 엔진).

출력: docs/prediction_log.json
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
DOCS = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS, "prediction_log.json")

WIN = 20    # 회귀 워크포워드 윈도우
WINF = 20   # 선물 워크포워드 윈도우


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


def build_futures_overnight(kdays, kospi, np, pd):
    """직전 KOSPI 마감(06:30 UTC) → 07:00 KST(=다음날 00:00 UTC-1.5h) 나스닥선물 야간수익률.
    ※ 07:00 채택: 백테스트상 05:30보다 MAE 낮음(밤사이 정보 더 반영, 여전히 개장 전).
    반환: {D: fov} — 07:00 고정 시점까지의 선물 움직임(개장 전 확정 정보)."""
    import yfinance as yf
    try:
        nq = yf.download("NQ=F", period="60d", interval="1h", progress=False)["Close"]
        if hasattr(nq, "columns"):       # CI yfinance 단일종목도 DataFrame → Series
            nq = nq.iloc[:, 0]
        nq = nq.dropna()
    except Exception as e:
        print(f"  [WARN] 선물(NQ=F) 다운로드 실패: {e}")
        return {}, None
    if len(nq) < 50:
        return {}, None
    nq.index = nq.index.tz_convert("UTC") if nq.index.tz else nq.index.tz_localize("UTC")
    nq = nq.sort_index()

    def ab(ts):  # ts 이전(같거나) 마지막 선물값
        sub = nq[nq.index <= ts]
        return float(sub.iloc[-1]) if len(sub) else None

    fov_map = {}
    for i in range(1, len(kdays)):
        D, P = kdays[i], kdays[i - 1]
        f0 = ab(pd.Timestamp(f"{P} 06:30", tz="UTC"))                       # 직전 KOSPI 마감(15:30 KST)
        f1 = ab(pd.Timestamp(f"{D} 00:00", tz="UTC") - pd.Timedelta(hours=2.0))  # 07:00 KST D
        if f0 and f1 and f0 > 0:
            r = f1 / f0 - 1
            if abs(r) < 0.2:
                fov_map[D] = r
    return fov_map, nq


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
    print("  KOSPI 개장 전 예측 vs 실제 추적기 (듀얼모델)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    kospi = naver_kospi_daily(95)
    if len(kospi) < WIN + 4:
        print("[ERROR] KOSPI 데이터 부족")
        return 1
    # 장중 미확정 바 제외 — 오늘은 마감(15:40 KST) 후에만 채점(실제 종가 오류 방지)
    _nowk = datetime.now(KST)
    if (_nowk.hour * 100 + _nowk.minute) < 1540 and _nowk.strftime("%Y-%m-%d") in kospi:
        del kospi[_nowk.strftime("%Y-%m-%d")]
    kdays = sorted(kospi.keys())

    # 전일종가·당일수익률 맵
    base_map, y_map = {}, {}
    for i in range(1, len(kdays)):
        D, P = kdays[i], kdays[i - 1]
        base_map[D] = kospi[P]
        y_map[D] = kospi[D] / kospi[P] - 1

    # 미국 SOX/NDX 일별 수익률 (회귀 입력)
    raw = yf.download(["^SOX", "^NDX"], period="5mo", interval="1d", progress=False, auto_adjust=True)
    cl = raw["Close"].dropna()
    sr = {d.strftime("%Y-%m-%d"): float(v) for d, v in cl["^SOX"].pct_change().items() if not np.isnan(v)}
    nr = {d.strftime("%Y-%m-%d"): float(v) for d, v in cl["^NDX"].pct_change().items() if not np.isnan(v)}

    # 회귀 표본 (D, P, base, y, s_prev, n_prev)
    reg = []
    for i in range(1, len(kdays)):
        D, P = kdays[i], kdays[i - 1]
        if P in sr and P in nr:
            reg.append((D, P, kospi[P], y_map[D], sr[P], nr[P]))

    def fit_beta(rows):
        Y = np.array([r[3] for r in rows])
        A = np.column_stack([[r[4] for r in rows], [r[5] for r in rows]])
        c, _, _, _ = np.linalg.lstsq(A, Y, rcond=None)
        return float(c[0]), float(c[1])

    # 회귀 워크포워드 예측
    reg_pred = {}  # D -> (pred, move, bS, bN, s, n)
    for j in range(WIN, len(reg)):
        D, P, base, y, s, n = reg[j]
        bS, bN = fit_beta(reg[j - WIN:j])
        move = bS * s + bN * n
        reg_pred[D] = (base * (1 + move), move, bS, bN, s, n)

    # 선물 야간 신호 + 워크포워드 예측
    fov_map, nq = build_futures_overnight(kdays, kospi, np, pd)
    fpairs = sorted((D, fov_map[D], y_map[D]) for D in fov_map if D in y_map)

    def fit_beta_f(rows):
        F = np.array([r[1] for r in rows])
        Y = np.array([r[2] for r in rows])
        return float((F @ Y) / (float(F @ F) or 1e-9))

    fut_pred = {}  # D -> (pred, beta_f, fov)
    for idx in range(len(fpairs)):
        if idx < WINF:
            continue
        D, fov, y = fpairs[idx]
        bf = fit_beta_f(fpairs[idx - WINF:idx])
        fut_pred[D] = (base_map[D] * (1 + bf * fov), bf, fov)

    # ── 채점 엔트리 조립 (마감 완료일만) ──
    def grade(base, pred, actual):
        err = (actual / pred - 1) * 100
        dir_ok = (pred >= base) == (actual >= base)
        return round(err, 2), round(abs(err), 2), bool(dir_ok)

    all_D = sorted(set(reg_pred) | set(fut_pred))
    entries = []
    for D in all_D:
        base = base_map[D]
        actual = kospi[D]
        e = {"date": D, "base": round(base, 2), "actual": round(actual, 2),
             "actual_pct": round(y_map[D] * 100, 2), "reg": None, "fut": None, "ens": None,
             "sox_ret": None, "ndx_ret": None, "fut_overnight": None}
        if D in reg_pred:
            pred, move, bS, bN, s, n = reg_pred[D]
            err, ae, dok = grade(base, pred, actual)
            e["reg"] = {"pred": round(pred, 0), "pct": round(move * 100, 2),
                        "err": err, "abs_err": ae, "dir": dok, "beta": {"sox": round(bS, 2), "ndx": round(bN, 2)}}
            e["sox_ret"], e["ndx_ret"] = round(s * 100, 2), round(n * 100, 2)
        if D in fut_pred:
            pred, bf, fov = fut_pred[D]
            err, ae, dok = grade(base, pred, actual)
            e["fut"] = {"pred": round(pred, 0), "pct": round((pred / base - 1) * 100, 2),
                        "err": err, "abs_err": ae, "dir": dok, "beta_f": round(bf, 2)}
            e["fut_overnight"] = round(fov * 100, 2)
        if e["reg"] and e["fut"]:
            pred = (e["reg"]["pred"] + e["fut"]["pred"]) / 2
            err, ae, dok = grade(base, pred, actual)
            e["ens"] = {"pred": round(pred, 0), "pct": round((pred / base - 1) * 100, 2),
                        "err": err, "abs_err": ae, "dir": dok}
        entries.append(e)

    entries = entries[-15:]

    def accuracy(key):
        vals = [e[key] for e in entries if e.get(key)]
        if not vals:
            return {"mae_pct": None, "dir_hit_rate": None, "n": 0}
        mae = round(sum(v["abs_err"] for v in vals) / len(vals), 2)
        dh = round(sum(1 for v in vals if v["dir"]) / len(vals) * 100)
        return {"mae_pct": mae, "dir_hit_rate": dh, "n": len(vals)}

    acc = {"reg": accuracy("reg"), "fut": accuracy("fut"), "ens": accuracy("ens")}

    # ── 다음(미마감) 세션 개장 전 고정 예측 — SK처럼 07:00 1회 고정·장중 수정 금지(persist freeze) ──
    now_k = datetime.now(KST)
    today_kst = now_k.strftime("%Y-%m-%d")
    hm = now_k.hour * 100 + now_k.minute
    LOCK_LO, LOCK_HI, CLOSE_HM = 700, 900, 1530  # 07:00~09:00 개장전 윈도우(07:00 신호), 15:30 마감
    L = kdays[-1]
    baseL = kospi[L]

    fresh = {"base": round(baseL, 2), "locked_at": now_k.strftime("%Y-%m-%d %H:%M KST"),
             "reg": None, "fut": None, "ens": None}
    if len(reg) >= WIN and L in sr and L in nr:
        bS, bN = fit_beta(reg[-WIN:])
        move = bS * sr[L] + bN * nr[L]
        fresh["reg"] = {"pred": round(baseL * (1 + move), 0), "pct": round(move * 100, 2),
                        "sox_ret": round(sr[L] * 100, 2), "ndx_ret": round(nr[L] * 100, 2)}
    if nq is not None and len(fpairs) >= WINF:
        try:
            sub = nq[nq.index <= pd.Timestamp.utcnow()]
            f0 = nq[nq.index <= pd.Timestamp(f"{L} 06:30", tz="UTC")]
            if len(sub) and len(f0) and float(f0.iloc[-1]) > 0:
                fov_now = float(sub.iloc[-1]) / float(f0.iloc[-1]) - 1
                bf = fit_beta_f(fpairs[-WINF:])
                fresh["fut"] = {"pred": round(baseL * (1 + bf * fov_now), 0),
                                "pct": round(bf * fov_now * 100, 2),
                                "overnight": round(fov_now * 100, 2)}
        except Exception as e:
            print(f"  [WARN] 선물 forward 예측 실패: {e}")
    if fresh["reg"] and fresh["fut"]:
        p = (fresh["reg"]["pred"] + fresh["fut"]["pred"]) / 2
        fresh["ens"] = {"pred": round(p, 0), "pct": round((p / baseL - 1) * 100, 2)}

    # 직전 출력 로드 → 장중 고정 유지(freeze)
    prev_today = {}
    try:
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            prev_today = (json.load(f) or {}).get("today") or {}
    except Exception:
        prev_today = {}

    if hm < CLOSE_HM:
        # 오늘 세션 대상 — 개장 전(07:00) 1회 고정 또는 장중 고정 유지(수정 금지)
        if prev_today.get("status") == "locked" and prev_today.get("target_date") == today_kst:
            today_block = prev_today
            print(f"  오늘({today_kst}) 예측 고정 유지 — 장중 수정 안 함")
        elif LOCK_LO <= hm < LOCK_HI:
            fresh.update({"status": "locked", "target_date": today_kst})
            today_block = fresh
        else:
            fresh.update({"status": "provisional", "target_date": today_kst})
            today_block = fresh
    else:
        # 마감 후 ~ 익일 새벽: 다음 거래일 잠정(내일 07:00 확정 예정)
        fresh.update({"status": "provisional", "target_date": "다음 거래일"})
        today_block = fresh
    target = today_block.get("target_date")

    # 이벤트 영향권 경고(FOMC/BOJ ±1거래일) — 점 예측은 못 고치므로 신뢰도만 명시(매 실행 갱신)
    try:
        from event_risk import event_risk_block
        _ev_date = today_kst if target == "다음 거래일" else target
        today_block = dict(today_block)
        today_block["event_risk"] = event_risk_block(_ev_date, window_days=1)
        if today_block["event_risk"]["flag"]:
            print(f"  ⚠️ 이벤트 영향권: {[e['title'] for e in today_block['event_risk']['events']]}")
    except Exception as e:
        print(f"  [WARN] event_risk 실패: {e}")

    # 반도체 디커플링 레짐 → 예측 신뢰도 경고(semi_decoupling.json, ≤1주기 지연 허용)
    try:
        with open(os.path.join(DOCS, "semi_decoupling.json"), encoding="utf-8") as f:
            _sd = (json.load(f) or {}).get("regime") or {}
        if isinstance(today_block, dict) and _sd.get("key") in ("high", "mid"):
            today_block["semi_regime"] = {"key": _sd["key"], "text": _sd.get("text")}
            print(f"  🔬 반도체 디커플링 레짐: {_sd['key']}")
    except Exception:
        pass

    for e in entries[-6:]:
        parts = []
        if e["reg"]:
            parts.append(f"회귀 {e['reg']['pred']:.0f}({e['reg']['err']:+.1f})")
        if e["fut"]:
            parts.append(f"선물 {e['fut']['pred']:.0f}({e['fut']['err']:+.1f})")
        print(f"  {e['date']} 실제 {e['actual']:.0f} | " + " · ".join(parts))
    print(f"\n  회귀 MAE ±{acc['reg']['mae_pct']}% 적중 {acc['reg']['dir_hit_rate']}% (n{acc['reg']['n']})")
    print(f"  선물 MAE ±{acc['fut']['mae_pct']}% 적중 {acc['fut']['dir_hit_rate']}% (n{acc['fut']['n']})")
    print(f"  앙상블 MAE ±{acc['ens']['mae_pct']}% 적중 {acc['ens']['dir_hit_rate']}% (n{acc['ens']['n']})")
    if today_block.get("ens") or today_block.get("fut") or today_block.get("reg"):
        print(f"  ▶ {target} 개장 전 고정 예측: "
              f"{(today_block.get('ens') or today_block.get('fut') or today_block.get('reg'))['pred']:.0f}")

    # ── 마감 후 자동 요약: 새로 채점된 날 1회 텔레그램 푸시(중복 방지) ──
    prev_notified = None
    try:
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            prev_notified = (json.load(f) or {}).get("last_notified")
    except Exception:
        pass
    last_notified = prev_notified
    try:
        from core import send_message, get_secret
    except Exception:
        send_message = get_secret = None
    if entries and send_message and get_secret and get_secret("TELEGRAM_FINANCE_BOT_TOKEN"):
        le = entries[-1]
        if le["date"] != prev_notified:
            def _seg(nm, m):
                return f"· {nm}: —" if not m else f"· {nm} {m['pred']:.0f} (오차 {m['err']:+.2f}% {'✓' if m['dir'] else '✗'})"
            msg = (f"📊 KOSPI 예측 채점 ({le['date']})\n"
                   f"실제 종가 {le['actual']:.0f} ({le['actual_pct']:+.2f}%)\n"
                   f"{_seg('선물', le.get('fut'))}\n{_seg('회귀', le.get('reg'))}\n{_seg('앙상블', le.get('ens'))}\n"
                   f"누적(최근 {acc['fut']['n']}일): 선물 MAE ±{acc['fut']['mae_pct']}% 방향 {acc['fut']['dir_hit_rate']}%")
            try:
                if send_message(msg):
                    last_notified = le["date"]
                    print(f"  텔레그램 채점 요약 발송: {le['date']}")
            except Exception as ex:
                print(f"  [WARN] 텔레그램 발송 실패: {ex}")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "models": "개장 전(07:00 KST) 고정 예측 — 회귀(미국 1일 시차) · 선물(NQ 야간) · 앙상블. 장중 보정 없음.",
        "window": WIN, "fut_window": WINF,
        "entries": entries,
        "accuracy": acc,
        "today": today_block,
        "last_notified": last_notified,
        "note": ("개장 전 확정 정보(미국 마감·나스닥선물 07:00)만으로 당일 종가 예측 → 마감 후 실제로 채점. "
                 "장중 시가갭 보정 같은 정보 누출은 제외. 워크포워드(해당일 제외)라 과적합 없음. 통계 추정."),
    }
    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} (채점 {len(entries)}건)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
