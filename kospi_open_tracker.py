"""KOSPI 시초가 예측 워크포워드 채점 — 매일 예측을 '개장 전에 고정'하고 실제 시가로 채점

왜 필요한가
----------
overseas_market_monitor 의 시초가 예측 가중치(선물 80% + EWY 20%)는 과거 데이터를
'뒤돌아보고' 정한 값이다(방향 적중 76.6%). 그 성적이 앞으로도 유지되는지는
**미리 기록해 두고 나중에 채점**해야만 알 수 있다. 이 파일이 그 역할을 한다.

동작
----
  1) 개장 전(KST 06~09시) 실행되면 그날의 예측을 로그에 **한 번만** 고정 기록한다.
     이미 기록이 있으면 덮어쓰지 않는다 — 장 시작 직전 값으로 갈아끼우면 성적이 부풀려진다.
  2) 아직 채점 안 된 과거 기록을 실제 시가로 채점한다.
       실제 갭 = 당일 시가 / 전일 종가 - 1
  3) 누적 성적(전체·최근 20·최근 60)을 계산해 저장한다.

데이터 출처
----------
  · 예측: docs/overseas_market.json 의 kospi_open_signal.composite_pct
  · 실제: yfinance ^KS11 일봉 (KRX 시가와 대조 검증 완료 — 아래 VALIDATION)
    KRX 는 GitHub IP 를 차단해 pykrx 를 쓸 수 없으므로 yfinance 를 쓴다.

VALIDATION (2026-08-03, pykrx 대조 · 최근 61거래일)
  시가 평균 차이 +0.0119% · 0.05% 초과 불일치 1일(2026-06-29, 0.72%)
  종가는 61일 전부 완전 일치(0.0000%)
  → 시가는 드물게 튀는 값이 있어 |갭| 15% 초과는 이상치로 보고 채점에서 제외한다.

출력: docs/kospi_open_log.json
🚨 정보·자기평가용 · 투자자문 아님.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS, "kospi_open_log.json")
SRC_FILE = os.path.join(DOCS, "overseas_market.json")

FREEZE_FROM, FREEZE_TO = 6, 9        # KST 06:00~08:59 에 그날 예측을 고정
MAX_KEEP = 250                       # 로그 보관 최대 일수
SANITY_GAP = 15.0                    # |갭| 15%p 초과 = 데이터 이상치로 간주(채점 제외)


def _load(path, default=None):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default if default is not None else {}


def _krx_open_day(d):
    """그날 KRX 가 여는가 (주말·공휴일 제외)."""
    try:
        from core.expiry import is_closed
        return not is_closed(d, "KR")
    except Exception:
        return d.weekday() < 5


def target_date(now):
    """이 예측이 향하는 거래일 — 오늘 개장 전이면 오늘, 아니면 다음 거래일."""
    d = now.date()
    if now.hour < 9 and _krx_open_day(d):
        return d
    for i in range(1, 12):
        nd = d + timedelta(days=i)
        if _krx_open_day(nd):
            return nd
    return None


def label_of(pct):
    """예측 문구 — overseas_market_monitor 의 구간과 동일하게 유지."""
    a = abs(pct)
    if a < 0.3:
        return "보합"
    if pct > 1.5:
        return "강한 갭상승"
    if pct > 0.5:
        return "갭상승"
    if pct < -1.5:
        return "강한 갭하락"
    if pct < -0.5:
        return "갭하락"
    return "약보합"


def fetch_actual_gaps(dates):
    """{날짜: 시초가 갭%} — yfinance ^KS11 일봉 기준."""
    if not dates:
        return {}
    try:
        import yfinance as yf
    except Exception as e:
        print(f"  [WARN] yfinance 없음: {e}")
        return {}
    try:
        df = yf.Ticker("^KS11").history(period="1y", interval="1d", auto_adjust=False)
    except Exception as e:
        print(f"  [WARN] ^KS11 다운로드 실패: {e}")
        return {}
    if df is None or len(df) < 2:
        return {}
    df = df.dropna(subset=["Open", "Close"])
    try:
        df.index = df.index.tz_localize(None)
    except Exception:
        pass
    out, keys = {}, [d.strftime("%Y-%m-%d") for d in df.index]
    for i in range(1, len(df)):
        ds = keys[i]
        if ds not in dates:
            continue
        prev_close = float(df["Close"].iloc[i - 1])
        if prev_close <= 0:
            continue
        gap = (float(df["Open"].iloc[i]) / prev_close - 1) * 100
        if abs(gap) > SANITY_GAP:                       # 데이터 이상치 방어(VALIDATION 참조)
            print(f"  [WARN] {ds} 갭 {gap:+.2f}% — 이상치로 채점 제외")
            continue
        out[ds] = round(gap, 3)
    return out


def stats_of(scored):
    """채점 완료 기록 → 성적(적중률·MAE·기준선 대비)."""
    if not scored:
        return None
    n = len(scored)
    hit = sum(1 for e in scored if e.get("hit")) / n * 100
    mae = sum(abs(e["error_pp"]) for e in scored) / n
    naive = sum(abs(e["actual_pct"]) for e in scored) / n     # '항상 보합' 기준선
    big = [e for e in scored if abs(e["predicted_pct"]) >= 0.5]   # 신호가 뚜렷했던 날만
    return {
        "n": n,
        "direction_hit_pct": round(hit, 1),
        "mae_pp": round(mae, 3),
        "naive_mae_pp": round(naive, 3),
        "beats_naive": bool(mae < naive),
        "strong_n": len(big),
        "strong_hit_pct": round(sum(1 for e in big if e.get("hit")) / len(big) * 100, 1) if big else None,
    }


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    log = _load(OUTPUT_FILE, {})
    entries = {e["date"]: e for e in (log.get("entries") or [])}

    # ── 1) 오늘(또는 다음 거래일) 예측 고정 기록 ──
    src = _load(SRC_FILE)
    sig = (src.get("kospi_open_signal") or {})
    comp = sig.get("composite_pct")
    td = target_date(now)
    if comp is None:
        print("[INFO] 예측 신호 없음 — 기록 스킵 (overseas_market.json 먼저 실행 필요)")
    elif td is None:
        print("[INFO] 대상 거래일을 찾지 못함 — 기록 스킵")
    elif str(td) in entries:
        print(f"[INFO] {td} 예측은 이미 고정됨({entries[str(td)]['predicted_pct']:+.3f}%) — 덮어쓰지 않음")
    elif not (FREEZE_FROM <= now.hour < FREEZE_TO):
        print(f"[INFO] 고정 시간대(KST {FREEZE_FROM}~{FREEZE_TO}시) 아님 (현재 {now.hour}시) — 기록 스킵")
    else:
        entries[str(td)] = {
            "date": str(td), "predicted_pct": round(float(comp), 3),
            "predicted_dir": "up" if comp > 0 else "down" if comp < 0 else "flat",
            "label": label_of(float(comp)),
            "ewy_pct": sig.get("ewy_pct"), "futures_pct": sig.get("futures_pct"),
            "recorded_at": now.strftime("%Y-%m-%d %H:%M KST"),
            "source_at": src.get("generated_at"),
            "actual_pct": None, "hit": None, "error_pp": None,
        }
        print(f"[고정] {td} 예측 {comp:+.3f}% ({label_of(float(comp))}) — 개장 전 기록 완료")

    # ── 2) 미채점 기록 채점 ──
    pending = [d for d, e in entries.items() if e.get("actual_pct") is None and d < str(now.date())]
    if pending:
        actual = fetch_actual_gaps(set(pending))
        newly = 0
        for d in pending:
            if d not in actual:
                continue
            e = entries[d]
            a, p = actual[d], e["predicted_pct"]
            e["actual_pct"] = a
            e["actual_dir"] = "up" if a > 0 else "down" if a < 0 else "flat"
            e["error_pp"] = round(p - a, 3)
            e["hit"] = bool((p > 0) == (a > 0)) if (p != 0 and a != 0) else None
            newly += 1
            print(f"  [채점] {d} 예측 {p:+.3f}% vs 실제 {a:+.3f}% → {'적중' if e['hit'] else '빗나감'}")
        print(f"[채점] {newly}건 완료 · 미채점 잔여 {len(pending) - newly}건")
    else:
        print("[채점] 채점 대상 없음")

    # ── 3) 성적 집계 ──
    ordered = sorted(entries.values(), key=lambda e: e["date"])[-MAX_KEEP:]
    scored = [e for e in ordered if e.get("actual_pct") is not None and e.get("hit") is not None]
    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "weights": {"ewy": 0.2, "us_futures": 0.8},
        "entries": ordered,
        "pending_count": sum(1 for e in ordered if e.get("actual_pct") is None),
        "stats_all": stats_of(scored),
        "stats_20": stats_of(scored[-20:]),
        "stats_60": stats_of(scored[-60:]),
        "backtest_reference": {"direction_hit_pct": 76.6, "mae_pp": 1.44, "n": 145,
                               "note": "과거 데이터로 가중치를 정할 때의 성적 — 아래 실시간 성적과 벌어지면 과최적화 신호"},
        "note": ("개장 전(KST 06~09시)에 예측을 고정 기록하고 다음 날 실제 시가로 채점하는 워크포워드 성적표. "
                 "실제 갭 = 당일 시가 ÷ 전일 종가 - 1 (yfinance ^KS11). "
                 "예측을 사후 수정하지 않으므로 백테스트와 달리 부풀려지지 않는다. "
                 "표본이 30건 미만이면 성적을 신뢰하기 어렵다. 투자자문 아님."),
    }
    if scored:
        s = out["stats_all"]
        print(f"\n[성적] 전체 {s['n']}건 · 방향 적중 {s['direction_hit_pct']}% · "
              f"MAE {s['mae_pp']}%p (기준선 {s['naive_mae_pp']}%p, {'우세' if s['beats_naive'] else '열세'})")
    else:
        print("\n[성적] 채점된 기록이 아직 없습니다 — 다음 거래일부터 쌓입니다.")

    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE} (기록 {len(ordered)}건 · 미채점 {out['pending_count']}건)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
