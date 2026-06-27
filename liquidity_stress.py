"""
시스템 유동성 위험 게이지
==========================
펀드 환매·헐값매각·은행 신용축소 같은 '직접 데이터'는 비공개·지연이라
실시간 불가 → 그 위기를 선행/동행하는 공개 지표를 신호등으로 종합.

지표(FRED, 공개·무료):
  · NFCI       — 시카고연준 금융여건지수(105개 지표 종합), 양수=평균보다 긴축
  · STLFSI4    — 세인트루이스연준 금융스트레스지수, 양수=평균 이상 스트레스
  · BAMLH0A0HYM2 — 하이일드 스프레드(OAS), 신용경색 1순위 선행
  · VIXCLS     — VIX 변동성, 주식 공포

출력: docs/liquidity_stress.json
🚨 통계·지표 종합. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "liquidity_stress.json")

# (id, 이름, green 임계, red 임계, 설명) — 값 < green=안정, green~red=주의, > red=경계
INDICATORS = [
    ("NFCI", "시카고연준 금융여건(NFCI)", 0.0, 0.5, "양수=평균보다 긴축(스트레스)"),
    ("STLFSI4", "세인트루이스 금융스트레스", 0.0, 1.0, "양수=평균 이상 스트레스"),
    ("BAMLH0A0HYM2", "하이일드 스프레드(OAS)", 4.0, 6.0, "신용경색 1순위 선행 (%)"),
    ("VIXCLS", "VIX 변동성", 20.0, 30.0, "주식 공포"),
]


def _fallback(sid):
    """FRED 타임아웃 대비 대체 소스 — VIX→yfinance, HY OAS→credit_spread.json. 전부 실패해 파일 미갱신되는 사고 방지."""
    try:
        if sid == "VIXCLS":
            import yfinance as yf
            import warnings
            warnings.filterwarnings("ignore")
            s = yf.Ticker("^VIX").history(period="3mo")["Close"].dropna()
            return [{"date": str(d.date()), "value": float(v)} for d, v in s.items()]
        if sid == "BAMLH0A0HYM2":
            with open(os.path.join(BASE_DIR, "docs", "credit_spread.json"), encoding="utf-8") as f:
                cs = json.load(f)
            for rr in cs.get("results", []):
                if (rr.get("id") == "BAMLH0A0HYM2" or "하이일드" in str(rr.get("name", ""))) and rr.get("latest_value") is not None:
                    return [{"date": rr.get("latest_date") or "", "value": float(rr["latest_value"])}]
    except Exception as e:
        print(f"    [fallback 실패] {sid}: {e}")
    return None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    # 기존 신용스프레드 모니터의 검증된 FRED 페처 재사용
    try:
        from credit_spread_monitor import fetch_fred_series
    except Exception as e:
        print(f"[ERROR] FRED 페처 임포트 실패: {e}")
        return 1

    print("=" * 55)
    print("  시스템 유동성 위험 게이지")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    since = (date.today() - timedelta(days=90)).strftime("%Y-%m-%d")
    items = []
    reds = ambers = 0
    for sid, name, g, r, desc in INDICATORS:
        rows = fetch_fred_series(sid, since=since)
        if not rows:
            rows = _fallback(sid)            # FRED 타임아웃 대비 대체 소스
            if rows:
                print(f"  [대체] {sid} ← 대체 소스(yfinance/credit_spread) 사용")
        if not rows:
            print(f"  [SKIP] {sid} 데이터 없음")
            continue
        rows.sort(key=lambda x: x["date"])
        latest = rows[-1]
        val = latest["value"]
        # 30일 전 대비 변화 (근사: ~22 영업일 전)
        prev = rows[-23] if len(rows) >= 23 else rows[0]
        chg = round(val - prev["value"], 3)
        if val > r:
            status, scol = "경계", "red"
            reds += 1
        elif val > g:
            status, scol = "주의", "amber"
            ambers += 1
        else:
            status, scol = "안정", "green"
        items.append({
            "id": sid, "name": name, "value": round(val, 2), "asof": latest["date"],
            "chg_30d": chg, "status": status, "color": scol,
            "green": g, "red": r, "desc": desc,
        })
        print(f"  {name}: {val:.2f} ({status}) 30d {chg:+.2f}")

    if not items:
        print("[ERROR] 지표 수집 실패")
        return 1

    # 종합 신호등
    if reds >= 2:
        overall, ocol = "🔴 경계 — 시스템 유동성 위험 고조", "red"
    elif reds >= 1 or ambers >= 2:
        overall, ocol = "🟡 주의 — 스트레스 징후", "amber"
    else:
        overall, ocol = "🟢 안정 — 유동성 양호", "green"
    print(f"\n종합: {overall} (경계 {reds} · 주의 {ambers})")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "overall": overall, "overall_color": ocol,
        "red_count": reds, "amber_count": ambers,
        "indicators": items,
        "note": ("펀드 환매·헐값매각·은행 신용축소의 직접 데이터는 비공개·지연이라 실시간 불가. "
                 "대신 연준 금융스트레스지수·신용스프레드·VIX로 같은 위기를 선행 포착. "
                 "지표 종합이며 투자 결정 단독 사용 금지."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
