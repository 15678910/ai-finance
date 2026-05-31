"""
경제 이벤트 캘린더 자동 수집
==============================
다음 소스에서 이벤트를 자동 수집합니다:
  1. FRED 릴리스 캘린더 API → CPI/PPI/GDP/고용 등 공식 발표일
  2. 하드코딩된 2026 FOMC/ECB/BOJ 일정 (연간 사전 공표)
  3. 수동 큐레이션 이벤트 (IPO, 지정학 등 뉴스 기반)

출력: docs/economic_calendar.json
"""

import json
import os
import sys
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta, date

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "economic_calendar.json")

# ── 1. 2026 중앙은행 일정 (연초 공표 기준) ─────────────────────────
CENTRAL_BANK_EVENTS = [
    # FOMC (미국)
    {"date": "2026-01-28", "title": "FOMC 금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리"]},
    {"date": "2026-03-18", "title": "FOMC 금리 결정 + 경제전망(SEP)", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리", "SEP"]},
    {"date": "2026-05-06", "title": "FOMC 금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리"]},
    {"date": "2026-06-17", "title": "FOMC 통화정책회의 (1일차)", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리"]},
    {"date": "2026-06-18", "title": "FOMC 금리 결정 + Kevin Warsh 기자회견", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "detail": "신임 의장 Kevin Warsh의 공식 포워드 가이던스 주목. SEP(점도표) 동반 발표.", "tags": ["FOMC", "Fed", "Warsh", "금리", "기자회견"]},
    {"date": "2026-07-29", "title": "FOMC 금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리"]},
    {"date": "2026-09-16", "title": "FOMC 금리 결정 + SEP", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리", "SEP"]},
    {"date": "2026-10-28", "title": "FOMC 금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리"]},
    {"date": "2026-12-09", "title": "FOMC 금리 결정 + SEP", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리", "SEP"]},
    # ECB (유럽)
    {"date": "2026-01-30", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리", "유로"]},
    {"date": "2026-03-06", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    {"date": "2026-04-17", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    {"date": "2026-06-11", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "detail": "유로존 경기 둔화 속 추가 인하 여부 결정.", "tags": ["ECB", "금리", "유로"]},
    {"date": "2026-07-17", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    {"date": "2026-09-11", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    {"date": "2026-10-23", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    {"date": "2026-12-11", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    # BOJ (일본)
    {"date": "2026-03-19", "title": "BOJ 금융정책결정회의", "category": "중앙은행", "impact": "HIGH", "region": "🇯🇵", "tags": ["BOJ", "엔화", "금리"]},
    {"date": "2026-05-01", "title": "BOJ 금융정책결정회의", "category": "중앙은행", "impact": "HIGH", "region": "🇯🇵", "tags": ["BOJ", "엔화"]},
    {"date": "2026-06-16", "title": "BOJ 금융정책결정회의 + 엔 캐리 청산 시그널", "category": "중앙은행", "impact": "HIGH", "region": "🇯🇵", "detail": "BOJ 추가 금리 인상 가능성. 엔 캐리 트레이드 청산 글로벌 리스크오프 트리거 주의.", "tags": ["BOJ", "엔화", "캐리트레이드", "리스크오프"]},
    {"date": "2026-07-30", "title": "BOJ 금융정책결정회의", "category": "중앙은행", "impact": "HIGH", "region": "🇯🇵", "tags": ["BOJ", "엔화"]},
    {"date": "2026-09-18", "title": "BOJ 금융정책결정회의", "category": "중앙은행", "impact": "HIGH", "region": "🇯🇵", "tags": ["BOJ", "엔화"]},
    # 한국은행
    {"date": "2026-01-16", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷", "tags": ["BOK", "한국은행", "기준금리"]},
    {"date": "2026-02-27", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷", "tags": ["BOK", "한국은행"]},
    {"date": "2026-04-17", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷", "tags": ["BOK", "한국은행"]},
    {"date": "2026-05-29", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷", "tags": ["BOK", "한국은행"]},
    {"date": "2026-07-10", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷", "tags": ["BOK", "한국은행"]},
]

# ── 2. 수동 큐레이션 (IPO/지정학 등 — 뉴스 기반) ───────────────────
CURATED_EVENTS = [
    # ── 젠슨 황(NVIDIA CEO) 방한 — 2026-06-05 ──────────────────────
    {
        "date": "2026-06-05",
        "title": "젠슨 황(NVIDIA CEO) 방한 — 2차 깐부 서밋",
        "category": "VIP방한",
        "impact": "HIGH",
        "region": "🇰🇷",
        "detail": (
            "COMPUTEX 2026(타이베이) 기조연설 직후 서울 방문. "
            "주요 면담: ① 최태원(SK그룹 회장) — HBM4 공급 확대·iHBM 냉각 기술, "
            "② 구광모(LG그룹 회장, 첫 만남) — Isaac 로보틱스 플랫폼 × LG CLOiD 협력, "
            "③ 이해진(네이버 창업자) — 옴니버스 디지털트윈 확장, "
            "④ 정의선(현대차그룹 회장) — Atlas 휴머노이드 자동화 (일정 확정 중). "
            "이재용(삼성 회장)은 해외 출장으로 불참. "
            "핵심 의제: Physical AI·로보틱스·소버린 AI(260,000 GPU 공급 2030년 목표). "
            "영향 종목: LG전자(+29.9%↑), 네이버(+14.1%↑), 삼성SDS(+20.3%↑), "
            "현대오토에버(+24.8%↑), SK하이닉스(HBM4 70% 점유), 삼성전자(HBM4E 샘플 최초 출하)."
        ),
        "ticker": "NVDA",
        "tags": ["NVIDIA", "젠슨황", "Physical AI", "HBM4", "로보틱스", "LG", "SK하이닉스", "방한"],
        "meetings": [
            {"name": "최태원", "role": "SK그룹 회장", "topic": "HBM4 공급·iHBM 냉각기술", "confirmed": True},
            {"name": "구광모", "role": "LG그룹 회장", "topic": "Isaac 로보틱스 × CLOiD 협력", "confirmed": True},
            {"name": "이해진", "role": "네이버 창업자", "topic": "Omniverse 디지털트윈", "confirmed": True},
            {"name": "정의선", "role": "현대차그룹 회장", "topic": "Atlas 휴머노이드 자동화", "confirmed": False},
            {"name": "이재용", "role": "삼성전자 회장", "topic": "HBM4E / 파운드리", "confirmed": False, "note": "해외 출장으로 불참"},
        ],
        "affected_stocks": [
            {"ticker": "066570", "name": "LG전자", "move": "+29.9%", "reason": "Isaac 로보틱스 협력"},
            {"ticker": "000660", "name": "SK하이닉스", "move": "HBM4 70%↑", "reason": "HBM4 최대 공급사"},
            {"ticker": "035420", "name": "네이버", "move": "+14.1%", "reason": "Omniverse 디지털트윈"},
            {"ticker": "018260", "name": "삼성SDS", "move": "+20.3%", "reason": "AI 인프라"},
            {"ticker": "005930", "name": "삼성전자", "move": "HBM4E 최초 출하", "reason": "HBM4E 샘플 세계 최초"},
        ],
    },
    {"date": "2026-06-10", "title": "SpaceX IPO", "category": "IPO", "impact": "HIGH", "region": "🇺🇸",
     "detail": "공모액 기준 역대 최대 규모 예상. 나스닥 상장, 기술주·방산우주 섹터 전반에 영향.", "tags": ["IPO", "우주항공", "기술"]},
    {"date": "2026-06-TBD", "title": "Anthropic IPO", "category": "IPO", "impact": "HIGH", "region": "🇺🇸",
     "detail": "AI 대형 언어모델 선두주자. 상장 시 AI 섹터 밸류에이션 기준점.", "tags": ["IPO", "AI", "LLM"]},
    {"date": "2026-06-TBD", "title": "OpenAI IPO / 구조 전환", "category": "IPO", "impact": "HIGH", "region": "🇺🇸",
     "detail": "비영리→영리 구조전환 후 상장 추진. ChatGPT 밸류에이션 $1000억+ 예상.", "tags": ["IPO", "AI", "ChatGPT"]},
    {"date": "2026-06-TBD", "title": "창신메모리(CXMT) A주 상장", "category": "IPO", "impact": "HIGH", "region": "🇨🇳",
     "detail": "중국 최대 DRAM 업체. 삼성·SK하이닉스 잠재 경쟁자. 한국 반도체 섹터 영향 주시.", "tags": ["IPO", "반도체", "DRAM", "중국"]},
    {"date": "2026-06-TBD", "title": "APEC 정상회담 (경주)", "category": "지정학", "impact": "MEDIUM", "region": "🇰🇷",
     "detail": "한국 경주 개최. 미중 양자회담 및 반도체·무역 협의 가능성.", "tags": ["APEC", "지정학", "미중", "무역"]},
]

# ── 3. FRED 릴리스 캘린더에서 경제지표 발표일 자동 수집 ──────────────
# FRED 경제 릴리스 ID 매핑
FRED_RELEASES = {
    10:  {"title": "미국 CPI (소비자물가)",       "category": "경제지표", "impact": "HIGH",   "region": "🇺🇸", "tags": ["CPI", "인플레이션", "Fed"]},
    31:  {"title": "미국 PPI (생산자물가)",        "category": "경제지표", "impact": "MEDIUM", "region": "🇺🇸", "tags": ["PPI", "인플레이션"]},
    50:  {"title": "미국 GDP (성장률)",            "category": "경제지표", "impact": "HIGH",   "region": "🇺🇸", "tags": ["GDP", "경기"]},
    51:  {"title": "미국 비농업고용 (NFP)",        "category": "경제지표", "impact": "HIGH",   "region": "🇺🇸", "tags": ["고용", "NFP", "Fed"]},
    103: {"title": "미국 소매판매",                "category": "경제지표", "impact": "MEDIUM", "region": "🇺🇸", "tags": ["소비", "소매"]},
    113: {"title": "미국 PCE 물가",               "category": "경제지표", "impact": "HIGH",   "region": "🇺🇸", "tags": ["PCE", "인플레이션", "Fed"]},
}


def fetch_fred_releases(api_key: str, start: str, end: str) -> list:
    """FRED 릴리스 캘린더 API에서 발표 일정을 수집합니다."""
    events = []
    for release_id, meta in FRED_RELEASES.items():
        try:
            params = urllib.parse.urlencode({
                "release_id": release_id,
                "realtime_start": start,
                "realtime_end": end,
                "file_type": "json",
                "api_key": api_key,
            })
            url = f"https://api.stlouisfed.org/fred/release/dates?{params}"
            req = urllib.request.Request(url, headers={"Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=10) as r:
                data = json.loads(r.read())
                for rd in data.get("release_dates", []):
                    events.append({
                        "date": rd["date"],
                        **meta,
                        "detail": f"BLS/BEA 공식 발표. 시장 컨센서스 대비 서프라이즈 시 변동성 확대.",
                    })
        except Exception as e:
            print(f"  [WARN] FRED release {release_id} 수집 실패: {e}")
    return events


def sort_events(events: list) -> list:
    def key(e):
        d = e.get("date", "")
        return "9999-99-99" if "TBD" in d else d
    return sorted(events, key=key)


def filter_upcoming(events: list, months_ahead: int = 3) -> list:
    """오늘 이후 ~ months_ahead개월 이내 이벤트만 반환."""
    today = date.today()
    cutoff = date(today.year + (today.month + months_ahead - 1) // 12,
                  (today.month + months_ahead - 1) % 12 + 1, 1)
    result = []
    for ev in events:
        d = ev.get("date", "")
        if "TBD" in d:
            result.append(ev)
            continue
        try:
            ev_date = date.fromisoformat(d)
            if today <= ev_date < cutoff:
                result.append(ev)
        except Exception:
            pass
    return result


def main():
    api_key = os.environ.get("FRED_API_KEY", "")
    now = datetime.now(KST)
    today = date.today()
    start_str = today.strftime("%Y-%m-%d")
    end_str = date(today.year + 1, today.month, 1).strftime("%Y-%m-%d")

    print("=" * 55)
    print("  경제 이벤트 캘린더 자동 수집")
    print(f"  KST: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    all_events = []

    # 중앙은행 일정 (하드코딩)
    cb_upcoming = filter_upcoming(CENTRAL_BANK_EVENTS)
    all_events.extend(cb_upcoming)
    print(f"[중앙은행] {len(cb_upcoming)}개 일정 로드")

    # FRED 경제지표 발표일 (API 자동 수집)
    if api_key:
        fred_events = fetch_fred_releases(api_key, start_str, end_str)
        fred_upcoming = filter_upcoming(fred_events)
        all_events.extend(fred_upcoming)
        print(f"[FRED] {len(fred_upcoming)}개 경제지표 발표일 수집")
    else:
        print("[WARN] FRED_API_KEY 없음 - 경제지표 발표일 자동 수집 생략")

    # 수동 큐레이션 (IPO/지정학)
    curated = filter_upcoming(CURATED_EVENTS)
    all_events.extend(curated)
    print(f"[큐레이션] {len(curated)}개 IPO/지정학 이벤트 로드")

    # 중복 제거 (같은 날짜+제목)
    seen = set()
    deduped = []
    for ev in all_events:
        key = (ev.get("date"), ev.get("title"))
        if key not in seen:
            seen.add(key)
            deduped.append(ev)

    deduped = sort_events(deduped)
    print(f"\n[합계] {len(deduped)}개 이벤트 (3개월 이내)")

    output = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "month": now.strftime("%Y-%m"),
        "note": "자동 수집 + 큐레이션 이벤트 (시뮬레이션/분석용. 투자 결정 금지)",
        "events": deduped,
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
