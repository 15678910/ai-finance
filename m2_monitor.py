"""
M2 증가율 모니터 (미국·한국)
============================
FRED API에서 US M2(M2SL)와 한국 M2(MYAGM2KRM189S)를 수집하여
전년동월비(YoY) 증가율을 계산합니다.

출력: docs/m2_data.json
"""

import json
import os
import sys
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "m2_data.json")

FRED_SERIES = {
    "us": {
        "id": "M2SL",
        "name": "미국 M2",
        "unit": "B USD",
        "color": "#22d3ee",
    },
    "kr": {
        "id": "MYAGM2KRM189S",
        "name": "한국 M2",
        "unit": "M KRW",
        "color": "#4ade80",
    },
}


def fetch_fred(series_id: str, api_key: str, months: int = 48) -> list:
    """FRED에서 월별 데이터를 가져옵니다."""
    from datetime import date, timedelta as td
    start = (date.today().replace(day=1) - td(days=months * 31)).strftime("%Y-%m-%d")
    params = urllib.parse.urlencode({
        "series_id": series_id,
        "observation_start": start,
        "file_type": "json",
        "sort_order": "asc",
        "api_key": api_key,
    })
    url = f"https://api.stlouisfed.org/fred/series/observations?{params}"
    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=15) as r:
            data = json.loads(r.read())
            obs = data.get("observations", [])
            # '.' 값 제거 (결측치)
            return [(o["date"], float(o["value"])) for o in obs if o["value"] != "."]
    except Exception as e:
        print(f"  [WARN] FRED {series_id} 수집 실패: {e}")
        return []


def calc_yoy(series: list) -> list:
    """전년동월비 증가율 계산. [(date, value, yoy_pct), ...]"""
    result = []
    date_to_val = {d: v for d, v in series}
    for date_str, val in series:
        try:
            year = int(date_str[:4])
            rest = date_str[4:]
            year_ago = f"{year - 1}{rest}"
            if year_ago in date_to_val and date_to_val[year_ago] > 0:
                yoy = (val - date_to_val[year_ago]) / date_to_val[year_ago] * 100
                result.append({
                    "date": date_str,
                    "value": round(val, 2),
                    "yoy_pct": round(yoy, 2),
                })
        except Exception:
            pass
    return result[-36:]  # 최근 36개월


def main():
    api_key = os.environ.get("FRED_API_KEY", "")
    if not api_key:
        print("[ERROR] FRED_API_KEY 환경변수 미설정")
        # Fallback: preserve existing file
        if os.path.exists(OUTPUT_FILE):
            print("[INFO] 기존 m2_data.json 유지")
            return 0
        sys.exit(1)

    print("=" * 50)
    print("  M2 증가율 모니터")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "series": {},
    }

    for key, meta in FRED_SERIES.items():
        print(f"\n[{meta['name']}] 수집 중...")
        raw = fetch_fred(meta["id"], api_key)
        if not raw:
            print(f"  데이터 없음")
            continue
        yoy_data = calc_yoy(raw)
        if yoy_data:
            latest = yoy_data[-1]
            print(f"  최신: {latest['date']}  YoY={latest['yoy_pct']:+.2f}%  값={latest['value']:,.0f}")
            output["series"][key] = {
                "name": meta["name"],
                "unit": meta["unit"],
                "color": meta["color"],
                "series_id": meta["id"],
                "data": yoy_data,
                "latest_yoy": latest["yoy_pct"],
                "latest_date": latest["date"],
                "latest_value": latest["value"],
            }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
