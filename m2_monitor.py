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
import urllib.error
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
    key = api_key.strip()   # 공백 제거
    params = urllib.parse.urlencode({
        "series_id": series_id,
        "api_key": key,
        "observation_start": start,
        "file_type": "json",
        "sort_order": "asc",
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
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")[:200]
        print(f"  [WARN] FRED {series_id} HTTP {e.code}: {body}")
        return []
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
BOK_BASE = "https://ecos.bok.or.kr/api"


def bok_get(_api_key: str, endpoint: str, timeout: int = 15) -> dict:
    """ECOS API 호출 헬퍼."""
    url = f"{BOK_BASE}/{endpoint}"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read())


def find_bok_m2_tables(api_key: str) -> tuple:
    """ECOS StatisticTableList에서 개정후/개정전 M2 통계표 코드 탐색.
    Returns: (after_candidates, before_candidates) — 각각 [(stat_code, item_code, label), ...]
    """
    keywords_after  = ["광의통화", "M2"]
    keywords_before = ["개정전", "구기준", "구잔액", "구방식"]
    after_cands, before_cands = [], []

    try:
        # 500개로 확장 (개정전 테이블이 200번 이후에 있을 수 있음)
        d = bok_get(api_key, f"StatisticTableList/{api_key}/json/kr/1/500/")
        tables = d.get("StatisticTableList", {}).get("row", [])
        print(f"  ECOS 통계표 {len(tables)}개 검색 중...")
        for t in tables:
            name = t.get("STAT_NAME", "")
            code = t.get("STAT_CODE", "")
            if not code:
                continue
            is_before = any(k in name for k in keywords_before)
            is_after  = any(k in name for k in keywords_after) and not is_before

            if is_before or is_after:
                print(f"  {'[개정전]' if is_before else '[개정후]'} 발견: {code} - {name}")
                # 항목 조회
                items = []
                try:
                    d2 = bok_get(api_key, f"StatisticItemList/{api_key}/json/kr/1/100/{code}/")
                    items = d2.get("StatisticItemList", {}).get("row", [])
                    print(f"    항목 {len(items)}개")
                except Exception as e:
                    print(f"    항목 조회 실패: {e}")

                kw_m2 = ["M2", "광의통화", "총계", "total"]
                matched = [(code, i.get("ITEM_CODE",""), i.get("ITEM_NAME",""))
                           for i in items if any(k in i.get("ITEM_NAME","") for k in kw_m2)]
                others  = [(code, i.get("ITEM_CODE",""), i.get("ITEM_NAME",""))
                           for i in items if not any(k in i.get("ITEM_NAME","") for k in kw_m2)][:3]
                if not items:
                    matched = [(code, "", "항목없음")]

                if is_before:
                    before_cands.extend(matched + others)
                else:
                    after_cands.extend(matched + others)
    except Exception as e:
        print(f"  [WARN] StatisticTableList 조회 실패: {e}")

    return after_cands, before_cands


def fetch_bok_m2(api_key: str) -> tuple:
    """한국은행 ECOS API에서 개정전/개정후 M2 데이터 수집.
    Returns: (rows_after, code_after, rows_before, code_before)
    """
    today = date.today()
    prev  = date(today.year, today.month, 1) - timedelta(days=1)
    start_ym = f"{today.year - 4}01"
    end_ym   = f"{prev.year}{prev.month:02d}"

    # 0단계: 키 유효성 (GDP 테스트)
    try:
        test_url = f"{BOK_BASE}/StatisticSearch/{api_key}/json/kr/1/1/722Y001/A/202301/202301/"
        req0 = urllib.request.Request(test_url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(req0, timeout=10) as r0:
            td = json.loads(r0.read())
        if "RESULT" in td and "StatisticSearch" not in td:
            res = td["RESULT"]
            print(f"  [키 검증] 오류: {res.get('CODE')} - {res.get('MESSAGE')}")
            if res.get("CODE") in ("API-100", "API-200", "API-300"):
                print("  API 키 유효하지 않음 — 스킵")
                return [], "", [], ""
        else:
            rows_t = td.get("StatisticSearch", {}).get("row", [])
            print(f"  [키 검증] 정상 ({len(rows_t)}행)")
    except Exception as e:
        print(f"  [키 검증] 네트워크 오류: {e}")

    # 1단계: 동적 탐색
    after_dynamic, before_dynamic = find_bok_m2_tables(api_key)

    # 2단계: 후보 목록 구성
    after_cands = list(after_dynamic) + [
        ("161Y005", "BBHS00", "M2 평잔 계절조정(검증됨)"),
        ("161Y006", "BBHA00", "M2 평잔 원계열(검증됨)"),
        ("161Y007", "BBGS00", "M2 말잔 계절조정"),
        ("161Y008", "BBGA00", "M2 말잔 원계열"),
    ]
    # 개정전 후보 — 1.7.x 섹션이 구기준(개정전) M2임
    # 1.1.x(개정후): 161Y005, 1.7.x(개정전): 101Y003 — 동일 구조, 다른 ECOS 섹션
    before_cands = list(before_dynamic) + [
        ("101Y003", "BBHS00", "1.7.x M2 계절조정(개정전 · 검증필요)"),
        ("101Y003", "",       "1.7.x M2(항목없음 시도)"),
        ("101Y004", "BBHA00", "1.7.x M2 원계열(개정전 추정)"),
        ("101Y004", "",       "1.7.x M2 원계열(항목없음)"),
        ("101Y001", "BBHS00", "1.7.x M2 말잔 계절조정(개정전 추정)"),
        ("161Y021", "BBHS00", "개정전 추정(161Y021)"),
        ("161Y022", "BBHS00", "개정전 추정(161Y022)"),
    ]

    # 개정전은 2025년 이후 병행공표 — 최근 날짜도 별도 시도
    recent_start = f"{today.year - 1}01"   # 1년 전부터

    def try_series(candidates: list, label: str, extra_start: "str | None" = None) -> tuple:
        """extra_start: 추가로 시도할 최근 시작날짜 (개정전용)."""
        date_ranges = [(start_ym, end_ym)]
        if extra_start:
            date_ranges.append((extra_start, end_ym))
        for code, icode, desc in candidates:
            suffix = f"/{icode}" if icode else ""
            # 먼저 StatisticItemList로 항목코드 확인 (항목코드 없을 때)
            if not icode:
                try:
                    d2 = bok_get(api_key, f"StatisticItemList/{api_key}/json/kr/1/20/{code}/")
                    items = d2.get("StatisticItemList", {}).get("row", [])
                    if items:
                        print(f"    [{code}] 항목 {len(items)}개: {[(i.get('ITEM_CODE',''), i.get('ITEM_NAME','')[:15]) for i in items[:4]]}")
                except Exception:
                    pass
            for s_ym, e_ym in date_ranges:
                url = f"{BOK_BASE}/StatisticSearch/{api_key}/json/kr/1/500/{code}/M/{s_ym}/{e_ym}{suffix}/"
                try:
                    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
                    with urllib.request.urlopen(req, timeout=15) as r:
                        d = json.loads(r.read())
                    if "RESULT" in d and "StatisticSearch" not in d:
                        res = d["RESULT"]
                        if s_ym == start_ym:  # 첫 시도만 출력
                            print(f"  [{label}/{code}{suffix}] {res.get('CODE')} - {res.get('MESSAGE')}")
                        continue
                    rows = d.get("StatisticSearch", {}).get("row", [])
                    print(f"  [{label}/{code}{suffix}] {desc} ({s_ym}-{e_ym}): {len(rows)}행 수신")
                    if rows:
                        print(f"    첫행: {rows[0]}")
                        return rows, code
                except Exception as e:
                    print(f"  [{label}/{code}] 수집 실패: {e}")
        return [], ""

    print("\n  [개정후] 탐색 중...")
    rows_after, code_after = try_series(after_cands, "개정후", None)
    print("\n  [개정전] 탐색 중... (최근 날짜도 시도)")
    rows_before, code_before = try_series(before_cands, "개정전", recent_start)
    return rows_after, code_after, rows_before, code_before


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
        rows_after, code_after, rows_before, code_before = fetch_bok_m2(bok_key)

        # 개정후 (post-revision, 현행 기준)
        kr_after_data = parse_bok_rows(rows_after, code_after) if rows_after else []
        if kr_after_data:
            latest = kr_after_data[-1]
            print(f"  [개정후] 최신: {latest['date']}  YoY={latest['yoy_pct']:+.2f}%")
            output["series"]["kr"] = {
                "name": "한국 M2 (개정후)", "unit": "전년동월비 %", "color": "#4ade80",
                "series_id": code_after, "data": kr_after_data,
                "latest_yoy": latest["yoy_pct"], "latest_date": latest["date"],
                "latest_value": latest["yoy_pct"],
                "revision": "after",
            }
        else:
            print("  [개정후] 데이터 없음")

        # 개정전 (pre-revision, 수익증권 포함 구기준)
        kr_before_data = parse_bok_rows(rows_before, code_before) if rows_before else []
        if kr_before_data:
            latest = kr_before_data[-1]
            print(f"  [개정전] 최신: {latest['date']}  YoY={latest['yoy_pct']:+.2f}%")
            output["series"]["kr_before"] = {
                "name": "한국 M2 (개정전)", "unit": "전년동월비 %", "color": "#f87171",
                "series_id": code_before, "data": kr_before_data,
                "latest_yoy": latest["yoy_pct"], "latest_date": latest["date"],
                "latest_value": latest["yoy_pct"],
                "revision": "before",
            }
        else:
            print("  [개정전] 데이터 없음 (개정전 시리즈 미발견)")
    else:
        print("\n[한국 M2] BOK_API_KEY 없음")
        print("  등록: https://ecos.bok.or.kr -> Open API -> API 키 발급")

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료 (시리즈 {len(output['series'])}개)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
