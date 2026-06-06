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
    # 6/17·6/18 FOMC는 CURATED_EVENTS의 Warsh 기자회견 항목으로 대체
    {"date": "2026-07-29", "title": "FOMC 금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸",
     "tags": ["FOMC", "Fed", "금리"],
     "impact_analysis": "Warsh 의장 체제 첫 하반기 회의. 9월 인하 여부 시그널 주목. 점도표(SEP) 미발표 회의."},
    {"date": "2026-09-16", "title": "FOMC 금리 결정 + SEP", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리", "SEP"]},
    {"date": "2026-10-28", "title": "FOMC 금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리"]},
    {"date": "2026-12-09", "title": "FOMC 금리 결정 + SEP", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸", "tags": ["FOMC", "Fed", "금리", "SEP"]},
    # ECB (유럽)
    {"date": "2026-01-30", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리", "유로"]},
    {"date": "2026-03-06", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    {"date": "2026-04-17", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    # 6/11 ECB는 CURATED_EVENTS로 대체
    {"date": "2026-07-17", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺",
     "tags": ["ECB", "금리"],
     "impact_analysis": "6월 인하 후 연속 인하 여부. 유로존 경기 회복 속도가 관건. EUR/USD 방향성 결정."},
    {"date": "2026-09-11", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    {"date": "2026-10-23", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    {"date": "2026-12-11", "title": "ECB 기준금리 결정", "category": "중앙은행", "impact": "HIGH", "region": "🇪🇺", "tags": ["ECB", "금리"]},
    # BOJ (일본)
    {"date": "2026-03-19", "title": "BOJ 금융정책결정회의", "category": "중앙은행", "impact": "HIGH", "region": "🇯🇵", "tags": ["BOJ", "엔화", "금리"]},
    {"date": "2026-05-01", "title": "BOJ 금융정책결정회의", "category": "중앙은행", "impact": "HIGH", "region": "🇯🇵", "tags": ["BOJ", "엔화"]},
    # 6/16 BOJ는 CURATED_EVENTS의 엔캐리 시그널 항목으로 대체
    {"date": "2026-07-30", "title": "BOJ 금융정책결정회의", "category": "중앙은행", "impact": "HIGH", "region": "🇯🇵",
     "tags": ["BOJ", "엔화"],
     "impact_analysis": "6월 인상 이후 연속 인상 여부. 1.25% 도달 시 엔캐리 청산 2차 파고 가능성."},
    {"date": "2026-09-18", "title": "BOJ 금융정책결정회의", "category": "중앙은행", "impact": "HIGH", "region": "🇯🇵", "tags": ["BOJ", "엔화"]},
    # 한국은행
    {"date": "2026-01-16", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷", "tags": ["BOK", "한국은행", "기준금리"]},
    {"date": "2026-02-27", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷", "tags": ["BOK", "한국은행"]},
    {"date": "2026-04-17", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷", "tags": ["BOK", "한국은행"]},
    {"date": "2026-05-29", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷", "tags": ["BOK", "한국은행"]},
    {"date": "2026-07-10", "title": "한국은행 금통위 기준금리 결정", "category": "중앙은행", "impact": "MEDIUM", "region": "🇰🇷",
     "tags": ["BOK", "한국은행"],
     "impact_analysis": "5월 금통위 매파 2명 인상 주장 → 7월 인상 기대. 인상 시 USD/KRW 하락, 은행주 수혜. 부동산 대출 규제와 연계."},
]

# ── 2. 수동 큐레이션 (IPO/지정학 등 — 뉴스 기반) ───────────────────
CURATED_EVENTS = [
    # ── 젠슨 황(NVIDIA CEO) 방한 ──────────────────────────────────────
    {
        "date": "2026-06-05",
        "title": "젠슨 황(NVIDIA CEO) 방한 — 2차 깐부 서밋",
        "category": "VIP방한",
        "impact": "HIGH",
        "region": "🇰🇷",
        "detail": (
            "COMPUTEX 2026(타이베이) 기조연설 직후 서울 방문. "
            "주제: Physical AI, 로보틱스(Isaac), 소버린 AI(2030년까지 260,000 GPU 공급 약속)."
        ),
        "ticker": "NVDA",
        "tags": ["NVIDIA", "젠슨황", "Physical AI", "HBM4", "로보틱스", "LG", "SK하이닉스"],
        "impact_analysis": (
            "📈 직접 수혜: LG전자(+29.9% 상한가·Isaac 로보틱스 협력), 네이버(+14.1%·Omniverse), "
            "삼성SDS(+20.3%), 현대오토에버(+24.8%). "
            "💾 반도체: SK하이닉스(HBM4 70% 점유·iHBM 냉각기술), 삼성전자(HBM4E 세계 최초 샘플 출하). "
            "⚠️ 주의: LG그룹 지주사·IT서비스 급등 — 방한 이후 차익 실현 가능성 고려."
        ),
        "meetings": [
            {"name": "최태원", "role": "SK그룹 회장", "topic": "HBM4 공급·iHBM 냉각기술", "confirmed": True},
            {"name": "구광모", "role": "LG그룹 회장", "topic": "Isaac 로보틱스 × LG CLOiD", "confirmed": True},
            {"name": "이해진", "role": "네이버 창업자", "topic": "Omniverse 디지털트윈", "confirmed": True},
            {"name": "정의선", "role": "현대차그룹 회장", "topic": "Atlas 휴머노이드 자동화", "confirmed": False},
            {"name": "이재용", "role": "삼성전자 회장", "topic": "HBM4E·파운드리", "confirmed": False, "note": "해외 출장 불참"},
        ],
        "affected_stocks": [
            {"ticker": "066570", "name": "LG전자", "move": "+29.9%", "reason": "Isaac 로보틱스 플랫폼 협력"},
            {"ticker": "000660", "name": "SK하이닉스", "move": "HBM4 70%↑", "reason": "HBM4 최대 공급사·iHBM"},
            {"ticker": "035420", "name": "네이버", "move": "+14.1%", "reason": "Omniverse 디지털트윈"},
            {"ticker": "018260", "name": "삼성SDS", "move": "+20.3%", "reason": "AI 인프라"},
            {"ticker": "005930", "name": "삼성전자", "move": "HBM4E 최초 출하", "reason": "Rubin Ultra용 HBM4E 선제 출하"},
        ],
    },
    # ── SpaceX IPO ────────────────────────────────────────────────
    {
        "date": "2026-06-12",
        "title": "SpaceX IPO (나스닥 상장)",
        "category": "IPO",
        "impact": "HIGH",
        "region": "🇺🇸",
        "detail": (
            "공모가격 결정 6/11, 나스닥 상장 6/12. 밸류에이션 $1.75조(~2,600조원) — 역대 최대 IPO. "
            "공모 규모 $750억. 주관: 골드만삭스·모건스탠리, 미래에셋증권(공동주관). "
            "매출 $185억(2025) — Starlink가 70% 이상 차지. 2026E 매출 $220-240억."
        ),
        "ticker": "SPCE",
        "tags": ["IPO", "SpaceX", "Starlink", "우주항공", "패시브펀드", "미래에셋"],
        "impact_analysis": (
            "🌐 글로벌 영향: S&P500 편입 시 패시브 펀드가 공모 물량의 ~19% 강제 매수 — "
            "이를 위해 기존 S&P500 종목(애플·엔비디아·마이크로소프트 등) 약 $9,500억 기계적 매도 예상. "
            "📈 직접 수혜 한국 기업: ① 미래에셋증권(공동주관, +190% YTD), "
            "② 미래에셋벤처투자(SpaceX 8,000억원 투자·+349%), "
            "③ 아주IB(미국법인 보유 SpaceX 지분·+446%), "
            "④ OCI홀딩스(폴리실리콘 공급·+173%), "
            "⑤ 스피어(특수합금 10년 계약), ⑥ 켄코아에어로스페이스(Tier-1 공급사). "
            "⚠️ 주의: 한화에어로스페이스·KAI·LIG넥스원은 직접 공급사 아님 — 테마 급등 후 차익실현 위험. "
            "💸 자본이동: 스페이스 ETF YTD +82%↑, 상장 후 Magnificent 7 기계적 매도 압력 주시."
        ),
    },
    # ── 미국 CPI ──────────────────────────────────────────────────
    {
        "date": "2026-06-11",
        "title": "미국 CPI 발표 (5월)",
        "category": "경제지표",
        "impact": "HIGH",
        "region": "🇺🇸",
        "detail": "BLS 소비자물가지수 5월 데이터 발표. 관세 인상의 물가 전이 여부 확인.",
        "tags": ["CPI", "인플레이션", "Fed", "Warsh"],
        "impact_analysis": (
            "📊 예상치 상회(HOT) 시: 달러 강세, 미국채 수익률 급등, KOSPI 외국인 이탈 압력. "
            "📊 예상치 하회(COOL) 시: Fed 9월 인하 기대 강화, 나스닥 랠리, 반도체·성장주 수혜. "
            "🔑 6/18 Warsh 기자회견 전 마지막 핵심 데이터 — 금리 경로 결정에 직결."
        ),
    },
    # ── ECB 기준금리 ───────────────────────────────────────────────
    {
        "date": "2026-06-11",
        "title": "ECB 기준금리 결정",
        "category": "중앙은행",
        "impact": "HIGH",
        "region": "🇪🇺",
        "detail": "유럽중앙은행 통화정책회의. 현재 3.25% — 유로존 성장 둔화 속 추가 인하 여부.",
        "tags": ["ECB", "금리", "유로", "라가르드"],
        "impact_analysis": (
            "✂️ 인하 시: EUR/USD 약세, 유럽 주식 상승, 글로벌 채권 강세 전반. "
            "한국: 달러 강세 → USD/KRW 상방 압력, 수출주(삼성·현대차) 단기 긍정. "
            "⏸️ 동결 시: ECB 신뢰도 훼손 우려, 유럽 경기 회의론 확산."
        ),
    },
    # ── BOJ 금융정책결정회의 + 엔 캐리 ────────────────────────────
    {
        "date": "2026-06-16",
        "title": "BOJ 금융정책결정회의 + 엔 캐리 청산 시그널",
        "category": "중앙은행",
        "impact": "HIGH",
        "region": "🇯🇵",
        "detail": (
            "현재 BOJ 정책금리 0.75%. OIS 시장 금리 인상 확률 74%(+25bp → 1.00%). "
            "4월 회의 6-3 동결(매파 3인 인상 주장), 전 심의위원 사쿠라이 '이번엔 올릴 것'. "
            "위험 요인: 도쿄 CPI 5월 +1.3%(목표 2% 하회, 6개월 연속 둔화) — BOJ 동결 빌미 가능."
        ),
        "tags": ["BOJ", "엔화", "캐리트레이드", "금리", "리스크오프"],
        "boj_hike_probability": 74,
        "carry_trade_size_bn": 400,
        "impact_analysis": (
            "🔴 인상 시(74% 확률): ① USD/JPY → 140 급락(현재 155-156), "
            "② 엔 캐리 잔고 $3,000-5,000억 청산 → VIX 급등(2024.8 사태: VIX 65, S&P -6% 3일). "
            "③ KOSPI 외국인 리밸런싱 매도(YTD +80% 고평가로 1순위 청산 대상). "
            "④ USD/KRW → 1,480-1,490 일시 하락(아시아 통화 강세 연동). "
            "⑤ 한국채 수익률 상승(BoK 인상 가속 기대). "
            "⑥ 삼성·SK하이닉스: 원화 강세 → 달러 매출 환산손 단기 부담. "
            "🟡 동결 시: 일시 안도 랠리, 하반기 인상 기대 유지 — 불확실성 지속."
        ),
    },
    # ── FOMC 기자회견 ─────────────────────────────────────────────
    {
        "date": "2026-06-18",
        "title": "FOMC + Kevin Warsh 기자회견",
        "category": "중앙은행",
        "impact": "HIGH",
        "region": "🇺🇸",
        "detail": (
            "신임 Fed 의장 Warsh 첫 공식 기자회견. SEP(점도표) 동반 발표. "
            "현재 FFR 3.75% — 시장은 9월 인하 기대 반영 중."
        ),
        "tags": ["FOMC", "Fed", "Warsh", "점도표", "금리"],
        "impact_analysis": (
            "🦅 매파 서프라이즈(인하 후퇴): 달러 급등, 나스닥 급락, KOSPI 외국인 이탈, "
            "USD/KRW 1,520+ 재진입 위험. "
            "🕊️ 비둘기(9월 인하 확인): 성장주 랠리, 반도체·배터리 강세, "
            "원화 강세(USD/KRW 1,450-1,480). "
            "🔑 BOJ 인상(6/16) + Warsh 매파(6/18) 동시 발생 시 — 복합 리스크오프 최대."
        ),
    },
    # ── Anthropic IPO ─────────────────────────────────────────────
    {
        "date": "2026-06-TBD",
        "title": "Anthropic IPO",
        "category": "IPO",
        "impact": "HIGH",
        "region": "🇺🇸",
        "detail": "Claude 개발사. 아마존·구글 투자. 밸류에이션 $400억+ 예상.",
        "tags": ["IPO", "AI", "Claude", "LLM"],
        "impact_analysis": (
            "🤖 AI 섹터 밸류에이션 기준점 형성 — OpenAI 다음 최대 LLM 업체. "
            "국내: 네이버(HyperCLOVA 비교), 카카오(KoGPT), 업스테이지 등 AI 기업 재평가 촉매."
        ),
    },
    # ── OpenAI ────────────────────────────────────────────────────
    {
        "date": "2026-06-TBD",
        "title": "OpenAI IPO / 구조 전환",
        "category": "IPO",
        "impact": "HIGH",
        "region": "🇺🇸",
        "detail": "비영리→영리법인 전환 후 상장. ChatGPT 밸류에이션 $1,500억+ 예상.",
        "tags": ["IPO", "AI", "ChatGPT", "OpenAI"],
        "impact_analysis": (
            "💡 AI 인프라 수요 재확인 — 삼성전자·SK하이닉스 HBM 수요 간접 증명. "
            "경쟁사 압박: 네이버·카카오 AI 투자 확대 불가피 → 비용 증가 우려."
        ),
    },
    # ── CXMT 상장 ─────────────────────────────────────────────────
    {
        "date": "2026-06-TBD",
        "title": "창신메모리(CXMT) A주 상장",
        "category": "IPO",
        "impact": "HIGH",
        "region": "🇨🇳",
        "detail": "중국 최대 DRAM 제조사. DDR5 양산 돌입. 미국 수출 제한 하에 내수 공략.",
        "tags": ["IPO", "반도체", "DRAM", "중국", "경쟁"],
        "impact_analysis": (
            "⚠️ 중장기 위협: 삼성·SK하이닉스 중국 내수 시장 잠식 가능성(현재 점유율 낮음). "
            "단기: DRAM 가격 하방 압력 우려로 SK하이닉스·삼성 밸류에이션 할인 요인. "
            "단, CXMT 기술력은 삼성 대비 2-3세대 격차 — HBM 경쟁은 아직 먼 미래."
        ),
    },
    # ── APEC ──────────────────────────────────────────────────────
    {
        "date": "2026-06-TBD",
        "title": "APEC 정상회담 (경주)",
        "category": "지정학",
        "impact": "MEDIUM",
        "region": "🇰🇷",
        "detail": "한국 경주 개최. 21개국 정상 참여. 미중 양자회담 가능성.",
        "tags": ["APEC", "지정학", "미중", "무역", "반도체"],
        "impact_analysis": (
            "🤝 미중 무역 합의 시: 한국 수출주(삼성·현대차·POSCO) 긍정, 원화 강세. "
            "반도체 수출 규제 완화 논의 가능성 → SK하이닉스 중국 HBM 판매 제한 해소 기대. "
            "⚡ 미중 갈등 격화 시: 한국의 진영 선택 압박 강화, 수출 불확실성 증가."
        ),
    },
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


# ── 발표 완료 실제값 수집 (FRED 시리즈) ───────────────────────────────
# 태그 키워드 → (series_id, units, 단위라벨, 소수자리)
#   units: pc1=전년동월비%, chg=전월대비변화, lin=수준값
FRED_VALUE_MAP = [
    (["CPI"],        "CPIAUCSL", "pc1", "% YoY", 1),   # 소비자물가 전년비
    (["PCE"],        "PCEPI",    "pc1", "% YoY", 1),   # PCE 물가 전년비
    (["PPI"],        "PPIFIS",   "pc1", "% YoY", 1),   # 생산자물가 전년비
    (["고용", "NFP"], "PAYEMS",   "chg", "천명",  0),   # 비농업고용 전월대비(천명)
    (["GDP"],        "A191RL1Q225SBEA", "lin", "%",  1),  # 실질GDP 성장률(연율)
    (["소매", "소비"], "RSAFS",    "pc1", "% YoY", 1),   # 소매판매 전년비
]


def fetch_indicator_value(api_key: str, series_id: str, units: str) -> tuple:
    """FRED에서 최신 발표값 + 직전값 반환. (latest_val, latest_date, prior_val) 또는 (None,None,None)."""
    try:
        params = urllib.parse.urlencode({
            "series_id": series_id, "api_key": api_key, "file_type": "json",
            "units": units, "sort_order": "desc", "limit": 3,
        })
        url = f"https://api.stlouisfed.org/fred/series/observations?{params}"
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=10) as r:
            data = json.loads(r.read())
        obs = [o for o in data.get("observations", []) if o.get("value") not in (".", "", None)]
        if not obs:
            return None, None, None
        latest = obs[0]
        prior = obs[1] if len(obs) > 1 else None
        return (float(latest["value"]), latest["date"],
                float(prior["value"]) if prior else None)
    except Exception as e:
        print(f"  [WARN] FRED 값 수집 실패 ({series_id}): {e}")
        return None, None, None


def enrich_events_with_actuals(events: list, api_key: str) -> list:
    """경제지표 이벤트에 발표완료 상태 + 실제값 부여.
    - 발표일 <= 오늘: status=released, 실제값/직전값/서프라이즈 첨부
    - 발표일 > 오늘:  status=upcoming
    """
    if not api_key:
        for ev in events:
            ev.setdefault("status", "upcoming")
        return events
    today = date.today()
    cache: dict = {}  # series_id → (val, date, prior)
    for ev in events:
        if ev.get("category") != "경제지표":
            ev.setdefault("status", "upcoming")
            continue
        d = ev.get("date", "")
        try:
            is_released = ("TBD" not in d) and (date.fromisoformat(d) <= today)
        except Exception:
            is_released = False
        ev["status"] = "released" if is_released else "upcoming"
        if not is_released:
            continue
        # 매핑 찾기
        tags = ev.get("tags", []) + [ev.get("title", "")]
        match = next((m for m in FRED_VALUE_MAP
                      if any(k in t for k in m[0] for t in tags)), None)
        if not match:
            continue
        _, sid, units, unit_label, dec = match
        if sid not in cache:
            cache[sid] = fetch_indicator_value(api_key, sid, units)
        val, vdate, prior = cache[sid]
        if val is None:
            continue
        ev["actual_value"] = round(val, dec)
        ev["actual_unit"] = unit_label
        ev["actual_period"] = (vdate or "")[:7]   # 기준월 (YYYY-MM)
        if prior is not None:
            ev["prior_value"] = round(prior, dec)
            diff = val - prior
            ev["surprise_vs_prior"] = round(diff, dec)
            # 물가·고용 상승 = 매파(인플레↑/경기과열), 하락 = 비둘기
            if abs(diff) < (0.1 if dec >= 1 else 1):
                ev["surprise_dir"] = "보합"
            elif diff > 0:
                ev["surprise_dir"] = "상승"
            else:
                ev["surprise_dir"] = "하락"
        print(f"  [발표완료] {ev['title']}: {ev['actual_value']}{unit_label} "
              f"(직전 {ev.get('prior_value','—')}, {ev.get('surprise_dir','')})")
    return events


def sort_events(events: list) -> list:
    def key(e):
        d = e.get("date", "")
        return "9999-99-99" if "TBD" in d else d
    return sorted(events, key=key)


def filter_window(events: list, months_ahead: int = 3, days_back: int = 14) -> list:
    """과거 days_back일 ~ 미래 months_ahead개월 이벤트 반환.
    과거(발표완료 결과 표시용)는 days_back일까지만 유지."""
    today = date.today()
    cutoff = date(today.year + (today.month + months_ahead - 1) // 12,
                  (today.month + months_ahead - 1) % 12 + 1, 1)
    back_limit = today - timedelta(days=days_back)
    result = []
    for ev in events:
        d = ev.get("date", "")
        if "TBD" in d:
            result.append(ev)
            continue
        try:
            ev_date = date.fromisoformat(d)
            if back_limit <= ev_date < cutoff:
                result.append(ev)
        except Exception:
            pass
    return result


# 하위호환 별칭
filter_upcoming = filter_window


def main():
    api_key = os.environ.get("FRED_API_KEY", "")
    now = datetime.now(KST)
    today = date.today()
    _ = today.strftime("%Y-%m-%d")  # start_str 예약
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

    # FRED 경제지표 발표일 (API 자동 수집) — 과거 14일분도 포함(발표완료 결과용)
    if api_key:
        past_str = (today - timedelta(days=14)).strftime("%Y-%m-%d")
        fred_events = fetch_fred_releases(api_key, past_str, end_str)
        fred_win = filter_window(fred_events)
        all_events.extend(fred_win)
        print(f"[FRED] {len(fred_win)}개 경제지표 발표일 수집 (과거 14일 포함)")
    else:
        print("[WARN] FRED_API_KEY 없음 - 경제지표 발표일 자동 수집 생략")

    # 수동 큐레이션 (IPO/지정학)
    curated = filter_upcoming(CURATED_EVENTS)
    all_events.extend(curated)
    print(f"[큐레이션] {len(curated)}개 IPO/지정학 이벤트 로드")

    # 중복 제거 (같은 날짜+제목) — 더 많은 필드를 가진 항목 우선
    seen: dict = {}  # key → index in deduped
    deduped = []
    for ev in all_events:
        key = (ev.get("date"), ev.get("title"))
        if key not in seen:
            seen[key] = len(deduped)
            deduped.append(ev)
        else:
            # 기존 항목보다 필드가 많으면(더 풍부하면) 교체
            existing_idx = seen[key]
            if len(ev) > len(deduped[existing_idx]):
                deduped[existing_idx] = ev

    # 발표완료 경제지표에 실제값 부여 (FRED)
    print("\n[발표완료 결과 수집]")
    deduped = enrich_events_with_actuals(deduped, api_key)

    deduped = sort_events(deduped)
    released = sum(1 for e in deduped if e.get("status") == "released")
    print(f"\n[합계] {len(deduped)}개 이벤트 (발표완료 {released}개)")

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
