"""
M2 증가율 모니터 (미국·한국)
============================
- 미국 M2: FRED JSON API (FRED_API_KEY 환경변수)
- 한국 M2: 한국은행 ECOS API (BOK_API_KEY 환경변수)

출력: docs/m2_data.json
"""

import json
import os
import sys
import urllib.request
import urllib.parse
from datetime import datetime, date, timezone, timedelta

KST      = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "m2_data.json")
USER_AGENT  = "Mozilla/5.0 (compatible; ai-finance-m2/1.0)"

# ── FRED: 미국 M2 ─────────────────────────────────────────────────────
FRED_SERIES = {
    "us": {"id": "M2SL", "name": "미국 M2", "unit": "B USD", "color": "#22d3ee"},
}


def fetch_fred(series_id: str, api_key: str) -> list:
    """FRED JSON API로 월별 M2 데이터 수집."""
    start = f"{date.today().year - 5}-01-01"
    params = urllib.parse.urlencode({
        "series_id": series_id,
        "api_key": api_key,
        "observation_start": start,
        "file_type": "json",
        "sort_order": "asc",
        "frequency": "m",
    })
    url = f"https://api.stlouisfed.org/fred/series/observations?{params}"
    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json",
                                                   "User-Agent": USER_AGENT})
        with urllib.request.urlopen(req, timeout=25) as r:
            d = json.loads(r.read())
        obs = d.get("observations", [])
        rows = [(o["date"], float(o["value"])) for o in obs if o["value"] not in (".", "")]
        print(f"  {series_id}: {len(rows)}개 월별 데이터 수집")
        return rows
    except Exception as e:
        print(f"  [WARN] FRED {series_id} 수집 실패: {e}")
        return []


def calc_yoy(series: list) -> list:
    """전년동월비(YoY) 증가율 계산."""
    date_to_val = {d: v for d, v in series}
    result = []
    for date_str, val in series:
        try:
            year_ago = f"{int(date_str[:4]) - 1}{date_str[4:]}"
            prev = date_to_val.get(year_ago)
            if prev and prev > 0:
                result.append({
                    "date": date_str,
                    "value": round(val, 2),
                    "yoy_pct": round((val - prev) / prev * 100, 2),
                })
        except Exception:
            pass
    return result[-36:]


# ── BOK ECOS: 한국 M2 ─────────────────────────────────────────────────
# ECOS M2 통계 코드 후보 (순서대로 시도)
BOK_M2_CODES = [
    ("BOBASE202Y", "M2(광의통화) 전년동월비"),
    ("101Y004",    "통화및유동성/M2"),
    ("101Y001",    "M2(원계열)"),
]
BOK_URL = "https://ecos.bok.or.kr/api/StatisticSearch/{key}/json/kr/1/500/{code}/M/{s}/{e}/"


def fetch_bok_m2(api_key: str) -> tuple:
    """한국은행 ECOS API에서 M2 데이터 수집. (rows, code_used) 반환."""
    today = date.today()
    start_ym = f"{today.year - 4}01"
    end_ym   = f"{today.year}{today.month:02d}"

    for code, label in BOK_M2_CODES:
        url = BOK_URL.format(key=api_key, code=code, s=start_ym, e=end_ym)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=15) as r:
                d = json.loads(r.read())

            # 오류 응답 처리
            if "RESULT" in d and "StatisticSearch" not in d:
                res = d["RESULT"]
                print(f"  [{code}] 오류: {res.get('CODE','?')} - {res.get('MESSAGE','?')}")
                continue

            inner = d.get("StatisticSearch", {})
            rows  = inner.get("row", [])
            print(f"  [{code}] {label}: {len(rows)}행 수신")
            if rows:
                print(f"    첫행 샘플: {rows[0]}")
                return rows, code
        except Exception as e:
            print(f"  [{code}] 수집 실패: {e}")

    return [], ""


def parse_bok_rows(rows: list, code: str = "") -> list:  # noqa: ARG001
    """ECOS 응답 행을 YoY 데이터로 변환.
    - 직접 전년동월비(%) 반환 시: DATA_VALUE를 yoy_pct로 사용
    - 잔액 데이터 시: calc_yoy로 별도 계산
    """
    raw = []
    for row in rows:
        ym      = row.get("TIME", "")
        val_str = row.get("DATA_VALUE", "").strip()
        if not ym or val_str in ("", "-", "N/A"):
            continue
        try:
            raw.append((f"{ym[:4]}-{ym[4:6]}-01", float(val_str)))
        except ValueError:
            pass
    if not raw:
        return []
    # 값 범위로 잔액/증가율 구분: 잔액은 수조~경 규모
    sample_val = abs(raw[-1][1]) if raw else 0
    if sample_val > 1_000_000:  # 잔액(원화 절대값)
        return calc_yoy(raw)
    else:  # 이미 % 증가율
        return [{"date": d, "value": v, "yoy_pct": v} for d, v in raw][-36:]


# ── 메인 ──────────────────────────────────────────────────────────────
def main():
    fred_key = os.environ.get("FRED_API_KEY", "")
    bok_key  = os.environ.get("BOK_API_KEY", "")

    print("=" * 52)
    print("  M2 증가율 모니터")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  FRED_API_KEY: {'있음' if fred_key else '없음'}")
    print(f"  BOK_API_KEY:  {'있음' if bok_key  else '없음'}")
    print("=" * 52)

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "series": {},
    }

    # 미국 M2
    if fred_key:
        for key, meta in FRED_SERIES.items():
            print(f"\n[{meta['name']}] 수집 중...")
            raw = fetch_fred(meta["id"], fred_key)
            yoy_data = calc_yoy(raw) if raw else []
            if yoy_data:
                latest = yoy_data[-1]
                print(f"  최신: {latest['date']}  YoY={latest['yoy_pct']:+.2f}%")
                output["series"][key] = {
                    "name": meta["name"], "unit": meta["unit"], "color": meta["color"],
                    "series_id": meta["id"], "data": yoy_data,
                    "latest_yoy": latest["yoy_pct"], "latest_date": latest["date"],
                    "latest_value": latest["value"],
                }
            else:
                print("  데이터 없음")
    else:
        print("\n[미국 M2] FRED_API_KEY 없음 - 스킵")

    # 한국 M2
    if bok_key:
        print("\n[한국 M2] 한국은행 ECOS API 수집 중...")
        rows, code_used = fetch_bok_m2(bok_key)
        kr_data = parse_bok_rows(rows, code_used) if rows else []
        if kr_data:
            latest = kr_data[-1]
            print(f"  최신: {latest['date']}  YoY={latest['yoy_pct']:+.2f}%")
            output["series"]["kr"] = {
                "name": "한국 M2", "unit": "전년동월비 %", "color": "#4ade80",
                "series_id": code_used, "data": kr_data,
                "latest_yoy": latest["yoy_pct"], "latest_date": latest["date"],
                "latest_value": latest["yoy_pct"],
            }
        else:
            print("  데이터 없음 (모든 통계코드 실패)")
    else:
        print("\n[한국 M2] BOK_API_KEY 없음")
        print("  등록: https://ecos.bok.or.kr → Open API → API 키 발급")

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료 (시리즈 {len(output['series'])}개)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
