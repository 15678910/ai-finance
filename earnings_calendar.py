"""
주요 일정 자동 캘린더 — 삼성·SK 등 핵심종목 실적발표일·배당락일 (yfinance 자동 수집)
===================================================================================
yfinance가 한국 상장사(.KS)의 차기 실적발표일·배당락일을 제공 → 매일 자동 갱신.
yfinance가 과거 날짜를 줄 때(지연)는 분기 주기로 다음 발표일을 추정(estimated 플래그).

출력: docs/key_events.json (D-day 정렬된 다가오는 일정)
🚨 발표일은 잠정치 — 회사 공식 공지·DART로 확정. 정보용·투자자문 아님.
"""

import json
import os
import sys
import warnings
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "key_events.json")

# 핵심 추적 종목 (이름, .KS 티커) — 삼성·SK 우선 + 대표 수혜 대형주
STOCKS = [
    ("삼성전자", "005930.KS"), ("SK하이닉스", "000660.KS"),
    ("한미반도체", "042700.KS"), ("두산로보틱스", "454910.KS"),
    ("HD현대일렉트릭", "267260.KS"), ("LS ELECTRIC", "010120.KS"),
    ("효성중공업", "298040.KS"),
]

# 잠정실적(가이던스) 오버레이 — yfinance는 '확정실적일'만 줘서 분기초 잠정실적을 놓침.
#   한국 대형주는 분기말 약 7영업일 후 잠정실적(영업이익 가이던스)을 먼저 공시.
#   과거가 되면 분기(3개월) 주기로 자동 롤포워드(추정 플래그). (date_anchor는 최근 확정 1건)
CURATED_PRELIM = [
    # (이름, 티커, 기준 잠정실적일, 비고)
    ("삼성전자", "005930.KS", "2026-07-07", "2분기 잠정실적(영업이익 가이던스) — 분기초 발표"),
]


def _add_months(d, months):
    m = d.month - 1 + months
    y = d.year + m // 12
    m = m % 12 + 1
    day = min(d.day, [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28,
                      31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
    return date(y, m, day)


def _roll_future(d, today):
    """과거 날짜면 분기(3개월) 단위로 미래까지 굴림 → (날짜, 추정여부)."""
    est = False
    guard = 0
    while d < today and guard < 8:
        d = _add_months(d, 3)
        est = True
        guard += 1
    return d, est


def _as_date(v):
    if v is None:
        return None
    if isinstance(v, list):
        v = v[0] if v else None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.fromisoformat(str(v)[:10]).date()
    except Exception:
        return None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    warnings.filterwarnings("ignore")

    try:
        import yfinance as yf
    except Exception as e:
        print(f"[ERROR] yfinance 임포트 실패: {e}")
        return 1

    today = datetime.now(KST).date()
    events = []
    for name, tk in STOCKS:
        try:
            cal = yf.Ticker(tk).calendar or {}
        except Exception as e:
            print(f"  [WARN] {name} calendar 실패: {e}")
            continue

        # 실적발표일 — yfinance가 주는 건 '확정실적'(전체 실적). 잠정실적은 아래 오버레이로 보완.
        ed = _as_date(cal.get("Earnings Date"))
        if ed:
            d2, est = _roll_future(ed, today)
            eps = cal.get("Earnings Average")
            events.append({
                "name": name, "ticker": tk, "type": "확정실적",
                "date": d2.isoformat(), "dday": (d2 - today).days, "estimated": est,
                "eps_est": round(eps) if isinstance(eps, (int, float)) else None,
            })
        # 배당락일
        xd = _as_date(cal.get("Ex-Dividend Date"))
        if xd:
            d2, est = _roll_future(xd, today)
            events.append({
                "name": name, "ticker": tk, "type": "배당락",
                "date": d2.isoformat(), "dday": (d2 - today).days, "estimated": est,
                "eps_est": None,
            })
        print(f"  {name}: 실적 {ed} · 배당락 {xd}")

    # 잠정실적(가이던스) 오버레이 — 분기초 발표. 과거면 분기 주기로 자동 롤포워드.
    for name, tk, anchor, memo in CURATED_PRELIM:
        ad = _as_date(anchor)
        if not ad:
            continue
        d2, est = _roll_future(ad, today)
        events.append({
            "name": name, "ticker": tk, "type": "잠정실적",
            "date": d2.isoformat(), "dday": (d2 - today).days, "estimated": est,
            "eps_est": None, "memo": memo,
        })
        print(f"  {name}: 잠정실적 {ad} → {d2}{' (추정)' if est else ''}")

    # 다가오는 순(D-day 오름차순). 과거(굴림 실패분)도 포함하되 뒤로.
    events.sort(key=lambda e: e["dday"])

    out = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "today": today.isoformat(),
        "events": events,
        "note": ("yfinance 자동 수집(.KS) — 핵심종목 차기 실적발표일·배당락일. "
                 "'추정'은 직전 일정 기준 분기 주기로 산출한 예상치(회사 공식 공지·DART로 확정). 정보용·투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}  (이벤트 {len(events)})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
