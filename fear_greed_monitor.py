"""
공포·탐욕지수 (Fear & Greed Index) — 미국증시(CNN) + 크립토(alternative.me)
==========================================================================
① 미국증시: CNN Fear & Greed Index (7개 세부지표 종합) — 봇 차단 있어 브라우저 헤더 필요.
② 크립토:   alternative.me Crypto Fear & Greed Index.
0=극단적 공포(역발상 매수기회), 100=극단적 탐욕(과열 주의). 앞서 만든 '과열 게이지'와 상호보완.

출력: docs/fear_greed.json
🚨 참고용(일 1회 갱신, 실시간 아님). 통계·심리 지표 · 투자자문 아님.
"""

import json
import os
import sys
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "fear_greed.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

CNN_URL = "https://production.dataviz.cnn.io/index/fearandgreed/graphdata"
CRYPTO_URL = "https://api.alternative.me/fng/?limit=60"

# CNN 세부지표(영문 키 → 한글)
CNN_COMPONENTS = [
    ("market_momentum_sp500", "시장 모멘텀"),
    ("stock_price_strength", "주가 강도"),
    ("stock_price_breadth", "주가 폭"),
    ("put_call_options", "풋/콜 옵션"),
    ("market_volatility_vix", "변동성(VIX)"),
    ("junk_bond_demand", "정크본드 수요"),
    ("safe_haven_demand", "안전자산 수요"),
]


def _get(url, headers):
    req = urllib.request.Request(url, headers=headers)
    return json.loads(urllib.request.urlopen(req, timeout=15).read().decode("utf-8", errors="replace"))


def fetch_cnn():
    try:
        d = _get(CNN_URL, {"User-Agent": UA, "Accept": "application/json",
                           "Referer": "https://edition.cnn.com/markets/fear-and-greed"})
        fg = d.get("fear_and_greed", {})
        hist = d.get("fear_and_greed_historical", {}).get("data", [])
        # 최근 ~60포인트로 다운샘플(스파크라인용)
        pts = hist[-90:]
        step = max(1, len(pts) // 60)
        series = [round(p["y"], 1) for i, p in enumerate(pts) if i % step == 0]
        comps = []
        for key, ko in CNN_COMPONENTS:
            c = d.get(key) or {}
            comps.append({"name": ko, "rating": c.get("rating")})
        return {
            "score": round(fg.get("score", 0), 1), "rating": fg.get("rating"),
            "prev_close": round(fg.get("previous_close", 0), 1),
            "week": round(fg.get("previous_1_week", 0), 1),
            "month": round(fg.get("previous_1_month", 0), 1),
            "year": round(fg.get("previous_1_year", 0), 1),
            "series": series, "components": comps,
        }
    except Exception as e:
        print(f"  [WARN] CNN 실패: {e}")
        return None


def fetch_crypto():
    try:
        d = _get(CRYPTO_URL, {"User-Agent": UA, "Accept": "application/json"})
        data = d.get("data", [])
        if not data:
            return None
        cur = data[0]
        series = [int(x["value"]) for x in reversed(data)]   # 오래된→최신
        return {
            "value": int(cur["value"]), "rating": cur.get("value_classification"),
            "yesterday": int(data[1]["value"]) if len(data) > 1 else None,
            "week": int(data[7]["value"]) if len(data) > 7 else None,
            "series": series,
        }
    except Exception as e:
        print(f"  [WARN] 크립토 실패: {e}")
        return None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    us = fetch_cnn()
    crypto = fetch_crypto()
    if not us and not crypto:
        print("[ERROR] 둘 다 수집 실패 — 기존 파일 보존.")
        return 1
    if us:
        print(f"  🇺🇸 미국증시 F&G: {us['score']} ({us['rating']}) · 1주전 {us['week']} · 1달전 {us['month']}")
    if crypto:
        print(f"  ₿ 크립토 F&G: {crypto['value']} ({crypto['rating']}) · 어제 {crypto['yesterday']}")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "us": us, "crypto": crypto,
        "note": ("공포·탐욕지수 — 미국증시=CNN(7개 세부지표 종합), 크립토=alternative.me. "
                 "0=극단적 공포(역발상 매수기회)·100=극단적 탐욕(과열 주의). 일 1회·투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
