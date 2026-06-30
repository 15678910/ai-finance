"""
KOSPI 변동성지수(VKOSPI) 실시간 모니터 — 한국판 VIX
=========================================================
VKOSPI = 코스피200 옵션 내재변동성 지수. 시장 공포·불확실성의 직접 척도.
네이버 실시간 API·yfinance엔 VKOSPI 심볼이 없어 → CNBC 차트 API(.KSVKOSPI)로 수집.

상태 밴드(통상): <20 안정 · 20~30 보통 · 30~40 경계 · >40 공포.
출력: docs/vkospi.json (최신값·일변화·상태·스파크라인 시계열·52주 범위)
🚨 정보 모니터링용. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "vkospi.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
CHART_URL = "https://ts-api.cnbc.com/harmony/app/charts/6M.json?symbol=.KSVKOSPI"


def _band(v):
    """VKOSPI 수준 → (상태, 색, 해석)."""
    if v is None:
        return "—", "var(--text-muted)", ""
    if v < 20:
        return "🟢 안정", "#4ade80", "변동성 낮음 — 시장 평온"
    if v < 30:
        return "🟡 보통", "#fbbf24", "통상 범위 — 경계 전 단계"
    if v < 40:
        return "🟠 경계", "#fb923c", "변동성 확대 — 위험회피 심리"
    return "🔴 공포", "#f87171", "극단적 변동성 — 패닉·급락 동반 구간"


def fetch_series():
    req = urllib.request.Request(CHART_URL, headers={"User-Agent": UA})
    raw = urllib.request.urlopen(req, timeout=15).read().decode("utf-8", "replace")
    bars = json.loads(raw).get("barData", {}).get("priceBars", [])
    out = []
    for b in bars:
        try:
            tt = str(b.get("tradeTime", ""))[:8]            # YYYYMMDD
            close = float(b.get("close"))
            if tt and close > 0:
                out.append({"t": f"{tt[:4]}-{tt[4:6]}-{tt[6:8]}", "close": round(close, 2)})
        except (TypeError, ValueError):
            continue
    return out


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    try:
        series = fetch_series()
    except Exception as e:
        print(f"[ERROR] VKOSPI 수집 실패: {e}")
        return 1
    if len(series) < 2:
        print("[ERROR] 시계열 부족")
        return 1

    latest = series[-1]
    prev = series[-2]
    value = latest["close"]
    chg = round(value - prev["close"], 2)
    chg_pct = round(chg / prev["close"] * 100, 2) if prev["close"] else None
    status, color, desc = _band(value)

    recent = series[-60:]                 # 스파크라인용 최근 60거래일
    win = [s["close"] for s in series[-252:]]   # 최근 52주
    hi52, lo52 = (round(max(win), 2), round(min(win), 2)) if win else (None, None)

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": latest["t"],
        "value": value, "change": chg, "change_pct": chg_pct,
        "status": status, "color": color, "desc": desc,
        "hi_52w": hi52, "lo_52w": lo52,
        "series": recent,
        "note": ("VKOSPI=코스피200 변동성지수(한국판 VIX). 출처 CNBC(.KSVKOSPI). "
                 "밴드: <20 안정·20~30 보통·30~40 경계·>40 공포. 정보용·투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"VKOSPI {value} ({chg:+.2f}, {chg_pct:+.2f}%) · {status} · asof {latest['t']} · 52w[{lo52}~{hi52}]")
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
