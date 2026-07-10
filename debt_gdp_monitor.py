"""
성장 대비 부채 비율 (정부부채/GDP) — 미국·한국·일본 비교
=========================================================
FRED에서 각국 '일반정부 총부채 / GDP(%)'를 수집해 3개국을 한 화면에서 비교.
  - 미국:  GGGDTAUSA188N (IMF 일반정부 총부채/GDP, 연간) — 없으면 GFDEGDQ188S(연방부채/GDP, 분기)로 폴백
  - 한국:  GGGDTAKRA188N (IMF 일반정부 총부채/GDP, 연간)
  - 일본:  GGGDTAJPA188N (IMF 일반정부 총부채/GDP, 연간)
'성장(GDP) 대비 부채'가 커질수록 재정 부담·금리·통화 리스크 증가. 일본은 세계 최고 수준(정책 여력 관건),
미국은 기축통화국 특권, 한국은 상대적으로 낮으나 증가 속도가 관전 포인트.

출력: docs/debt_gdp.json
🚨 참고용(FRED·IMF 공개데이터, 워크플로 갱신). 투자자문 아님.
"""

import json
import os
import sys
import urllib.parse
import urllib.request
from datetime import date, datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "debt_gdp.json")

# 국가별 (표시명, 국기, 색, FRED 시리즈 후보[우선순위순])
COUNTRIES = [
    ("US", "미국", "🇺🇸", "#60a5fa", ["GGGDTAUSA188N", "GFDEGDQ188S"]),
    ("KR", "한국", "🇰🇷", "#f472b6", ["GGGDTAKRA188N"]),
    ("JP", "일본", "🇯🇵", "#fbbf24", ["GGGDTAJPA188N"]),
]


def fred_observations(api_key, series_id, years_back=40):
    """FRED 원시 관측값 [(date, value), ...] 오름차순."""
    if not api_key:
        return []
    try:
        start = f"{date.today().year - years_back}-01-01"
        params = urllib.parse.urlencode({
            "series_id": series_id, "api_key": api_key.strip(),
            "observation_start": start, "file_type": "json", "sort_order": "asc",
        })
        url = f"https://api.stlouisfed.org/fred/series/observations?{params}"
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=15) as r:
            d = json.loads(r.read())
        return [(o["date"], float(o["value"])) for o in d.get("observations", [])
                if o.get("value") not in (".", "", None)]
    except Exception as e:
        print(f"  [WARN] FRED {series_id} 실패: {e}")
        return []


def annualize(obs):
    """(date, value) 리스트 → {year: 연말(마지막) 값} (분기·연간 모두 연 단위로 정규화)."""
    by_year = {}
    for d, v in obs:  # 오름차순이라 뒤가 최신 → 연내 마지막 관측이 남음
        try:
            y = int(str(d)[:4])
        except ValueError:
            continue
        by_year[y] = v
    return by_year


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    api_key = os.environ.get("FRED_API_KEY", "")
    if not api_key:
        try:
            from core import get_secret
            api_key = get_secret("FRED_API_KEY") or ""
        except Exception:
            pass
    if not api_key:
        print("[ERROR] FRED_API_KEY 없음 — 기존 파일 보존.")
        return 1

    now = datetime.now(KST)
    countries = []
    for code, name, flag, color, series_ids in COUNTRIES:
        obs, used = [], None
        for sid in series_ids:
            obs = fred_observations(api_key, sid)
            if obs:
                used = sid
                break
        if not obs:
            print(f"  [WARN] {name}: 데이터 없음(시도 {series_ids})")
            continue
        by_year = annualize(obs)
        series = [[y, round(by_year[y], 1)] for y in sorted(by_year)]
        latest_year = series[-1][0]
        latest = series[-1][1]
        countries.append({
            "code": code, "name": name, "flag": flag, "color": color,
            "series_id": used, "latest": latest, "latest_year": latest_year,
            "series": series,
        })
        print(f"  {flag} {name}: {latest}% ({latest_year}) · {len(series)}년치 [{used}]")

    if not countries:
        print("[ERROR] 수집 실패 — 기존 파일 보존.")
        return 1

    # 공통 연도 범위(차트 x축)
    all_years = sorted({p[0] for c in countries for p in c["series"]})
    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "countries": countries,
        "year_min": all_years[0], "year_max": all_years[-1],
        "note": ("각국 일반정부 총부채/GDP(%). 미국=IMF 일반정부(없으면 연방부채/GDP 분기), "
                 "한국·일본=IMF 일반정부(연간). 성장 대비 부채가 클수록 재정·금리·통화 리스크↑. "
                 "FRED/IMF 공개데이터 · 워크플로 갱신 · 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE}  ({len(countries)}개국)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
