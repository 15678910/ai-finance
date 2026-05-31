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
    # 한국 M2: FRED 시리즈 단종(2017). 한국은행 ECOS API(BOK_API_KEY) 사용 권장.
    # BOK_API_KEY 미설정 시 한국 M2 미표시.
}

# 한국은행 ECOS M2 시리즈 코드
BOK_M2_CODE  = "101Y004"   # M2(광의통화) 전년동월비
BOK_STAT_URL = "https://ecos.bok.or.kr/api/StatisticSearch/{key}/json/kr/1/500/{code}/M/{start}/{end}/"


FRED_CSV_BASE = "https://fred.stlouisfed.org/graph/fredgraph.csv"
USER_AGENT = "Mozilla/5.0 (compatible; ai-finance-m2-monitor/1.0)"


def fetch_fred(series_id: str) -> list:
    """FRED CSV 엔드포인트로 월별 데이터를 수집합니다 (API 키 불필요).
    credit_spread_monitor.py와 동일한 방식 사용.
    """
    from datetime import date
    # 4년치 데이터 확보 (YoY 계산에 13개월 이상 필요)
    start_year = date.today().year - 4
    start = f"{start_year}-01-01"
    url = f"{FRED_CSV_BASE}?id={series_id}&vintage_date={start}"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(req, timeout=20) as r:
            text = r.read().decode("utf-8")
        rows = []
        for line in text.strip().splitlines()[1:]:   # 헤더 skip
            parts = line.split(",")
            if len(parts) == 2 and parts[1].strip() not in (".", ""):
                try:
                    rows.append((parts[0].strip(), float(parts[1].strip())))
                except ValueError:
                    pass
        rows.sort(key=lambda x: x[0])
        print(f"  {series_id}: {len(rows)}개 월별 데이터 수집")
        return rows
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
    # CSV 방식은 API 키 불필요 (credit_spread_monitor.py와 동일)
    print("=" * 50)
    print("  M2 증가율 모니터")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "series": {},
    }

    # ── 미국 M2 (FRED CSV, 키 불필요) ────────────────────────────────
    for key, meta in FRED_SERIES.items():
        print(f"\n[{meta['name']}] 수집 중...")
        raw = fetch_fred(meta["id"])
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

    # ── 한국 M2 (한국은행 ECOS API, BOK_API_KEY 환경변수 필요) ─────────
    bok_key = os.environ.get("BOK_API_KEY", "")
    if bok_key:
        print("\n[한국 M2] 한국은행 ECOS API 수집 중...")
        try:
            from datetime import date
            start_ym = f"{date.today().year - 4}01"
            end_ym   = f"{date.today().year}{date.today().month:02d}"
            url = BOK_STAT_URL.format(key=bok_key, code=BOK_M2_CODE,
                                      start=start_ym, end=end_ym)
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=15) as r:
                d = json.loads(r.read())
            rows = d.get("StatisticSearch", {}).get("row", [])
            # ECOS는 YoY % 직접 반환 (전년동월비)
            kr_data = []
            for row in rows:
                ym = row.get("TIME", "")           # "202601" 형식
                val_str = row.get("DATA_VALUE", "")
                if ym and val_str not in ("", "-", "N/A"):
                    try:
                        date_str = f"{ym[:4]}-{ym[4:6]}-01"
                        yoy = float(val_str)
                        kr_data.append({"date": date_str, "value": yoy, "yoy_pct": yoy})
                    except ValueError:
                        pass
            kr_data = kr_data[-36:]
            if kr_data:
                latest = kr_data[-1]
                print(f"  최신: {latest['date']}  YoY={latest['yoy_pct']:+.2f}%")
                output["series"]["kr"] = {
                    "name": "한국 M2",
                    "unit": "전년동월비 %",
                    "color": "#4ade80",
                    "series_id": BOK_M2_CODE,
                    "data": kr_data,
                    "latest_yoy": latest["yoy_pct"],
                    "latest_date": latest["date"],
                    "latest_value": latest["yoy_pct"],
                }
        except Exception as e:
            print(f"  [WARN] 한국은행 ECOS 수집 실패: {e}")
    else:
        print("\n[한국 M2] BOK_API_KEY 미설정 - 한국은행 ECOS API 키 등록 후 활성화")
        print("  등록: https://ecos.bok.or.kr → 개발자서비스 → API키발급")

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
