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
    # (미국 CPI·NFP·PPI·GDP·PCE·소매판매는 Nasdaq 캘린더에서 자동 수집 — 수동 큐레이션 제거)
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
            "현재 BOJ 정책금리 0.75%. 이코노미스트 컨센서스 ~94% 인상 전망(+25bp → 1.00%, "
            "Reuters 폴 51명 중 49명 인상 · OIS 거의 확실 반영). 전 심의위원 사쿠라이 '이번엔 올릴 것'. "
            "위험 요인: 도쿄 CPI 둔화는 동결 빌미였으나 중동발 에너지 급등·엔 약세로 인상 쪽 우세. "
            "※ 수기 큐레이션(2026-06-15 기준) · 실시간 OIS 자동연동 아님."
        ),
        "tags": ["BOJ", "엔화", "캐리트레이드", "금리", "리스크오프"],
        "boj_hike_probability": 94,
        "carry_trade_size_bn": 400,
        "impact_analysis": (
            "🔴 인상 시(컨센서스 ~94%): ① USD/JPY 급락=엔 강세(현재 160 → 150선 테스트), "
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
    # ── SK하이닉스 미국 ADR 상장 (진행 중 catalyst) ───────────────
    {
        "date": "2026-07-10",
        "title": "SK하이닉스 미국 ADR 나스닥 상장 (확정)",
        "category": "기업",
        "impact": "HIGH",
        "region": "🇰🇷🇺🇸",
        "detail": ("신주 DR(유상증자) 확정(이사회 6/24) — 총 ₩45.45조·신주 약 17.79M주(약 2.44% 희석)·"
                   "발행가 2,555,000원(=6/23 폭락 종가). 나스닥 상장 7/10·청약/납입 7/14·신주(원주) 상장 7/29. "
                   "자금=HBM·AI 시설(capex). 주관 BofA·Citi·GS·JPM."),
        "tags": ["SK하이닉스", "ADR", "유상증자", "오버행", "청약", "나스닥", "HBM"],
        "impact_analysis": (
            "확정: 신주발행(유상증자=희석) ₩45.45조, 약 2.44% 희석. 발행가가 시장가 수준(2,555,000)이라 추가 할인 충격은 작음. "
            "🔴 단기 리스크=오버행·'역유입': 7/10 ADR 거래개시 후 ADR가격<원주면 차익전환 매도로 KOSPI 원주 하방압력. "
            "7/29 신주(원주) 상장이 희석분 유통 반영 분기점. "
            "🟢 중기 상방: 45조 capex가 HBM 성장으로 이어지고 나스닥 상장이 코리아 디스카운트 해소(한국 선행 PER 5.8배 vs 비한국 공급처 16.7배)로 작동 시. "
            "핵심 줄다리기 = 신주 2.44% 희석 vs 45조 capex 성장 + 재평가."
        ),
    },
    # ── 2026년 7월 둘째 주(7/6~7/10) 글로벌 증시 주요 일정 ─────────────
    {"date": "2026-07-06", "title": "달러-원 환율 24시간 거래 개시", "category": "정책", "impact": "MEDIUM", "region": "🇰🇷",
     "tags": ["원달러", "외환시장", "24시간"],
     "impact_analysis": "국내 외환시장 24시간 거래 개시 — 야간 역외 변동성이 국내 지표에 실시간 반영. 환율 급변 시 외국인 수급·수출주 변동성 확대 가능."},
    {"date": "2026-07-06", "title": "국제 머신러닝 학회(ICML)", "category": "기술", "impact": "LOW", "region": "🌍", "tags": ["ICML", "AI", "머신러닝"]},
    {"date": "2026-07-07", "title": "나토(NATO) 정상회의", "category": "지정학", "impact": "MEDIUM", "region": "🌍",
     "tags": ["NATO", "지정학", "국방"],
     "impact_analysis": "국방비 증액·우크라이나 지원 논의. 방산주(한화에어로·현대로템·LIG넥스원) 수출 모멘텀과 연계."},
    {"date": "2026-07-07", "title": "캐나다 초계 잠수함 프로젝트(CPSP) 최종 사업자 발표", "category": "지정학", "impact": "MEDIUM", "region": "🇨🇦🇰🇷",
     "tags": ["방산", "잠수함", "한화오션", "HD현대중공업"],
     "impact_analysis": "한화오션·HD현대중공업이 유력 후보. 최대 12척(수십조원) 대형 수주 — 선정 시 해당 조선·방산주 강한 상방 촉매."},
    {"date": "2026-07-07", "title": "美 대체 관세 부과 공청회", "category": "통상", "impact": "MEDIUM", "region": "🇺🇸", "tags": ["관세", "통상", "무역"]},
    {"date": "2026-07-07", "title": "삼성전자 2026 2분기 잠정실적 발표", "category": "실적", "impact": "HIGH", "region": "🇰🇷", "ticker": "005930",
     "tags": ["삼성전자", "잠정실적", "반도체", "HBM"],
     "impact_analysis": "잠정치는 매출·영업이익 총액만 공개(부문별은 7/29 확정). HBM·DRAM 회복 강도와 파운드리 적자 축소 여부가 관전 포인트. 컨센서스 상회 시 반도체 섹터 전반 리레이팅."},
    {"date": "2026-07-07", "title": "LG전자 2026 2분기 잠정실적 발표", "category": "실적", "impact": "MEDIUM", "region": "🇰🇷", "ticker": "066570", "tags": ["LG전자", "잠정실적", "가전"]},
    {"date": "2026-07-07", "title": "스페이스X 나스닥100 편입(방송 일정 기준)", "category": "지수", "impact": "MEDIUM", "region": "🇺🇸", "tags": ["SpaceX", "나스닥100", "지수편입"]},
    {"date": "2026-07-08", "title": "美 FOMC 의사록 공개(6월 회의)", "category": "중앙은행", "impact": "HIGH", "region": "🇺🇸",
     "tags": ["FOMC", "의사록", "Fed", "금리"],
     "impact_analysis": "6월 회의 내부 논의 공개 — 9월 인하 시그널·위원 간 이견 확인. 매파적이면 나스닥·반도체 조정, 비둘기면 위험자산 반등."},
    {"date": "2026-07-08", "title": "韓 5월 경상수지", "category": "경제지표", "impact": "LOW", "region": "🇰🇷", "tags": ["경상수지", "무역"]},
    {"date": "2026-07-09", "title": "韓 옵션 만기일", "category": "파생만기", "impact": "MEDIUM", "region": "🇰🇷",
     "tags": ["옵션만기", "파생", "수급"],
     "impact_analysis": "월물 옵션 만기 — 만기 전후 프로그램 매매·변동성 확대 주의. VKOSPI와 함께 확인."},
    {"date": "2026-07-09", "title": "이재명 대통령 몽골 국빈방문", "category": "지정학", "impact": "LOW", "region": "🇰🇷🇲🇳", "tags": ["외교", "자원", "몽골"]},
    {"date": "2026-07-09", "title": "中 6월 CPI", "category": "경제지표", "impact": "MEDIUM", "region": "🇨🇳",
     "tags": ["CPI", "물가", "중국", "디플레"],
     "impact_analysis": "중국 디플레 우려 지속 여부 확인. 약세 지속 시 소재·화학·중국소비 관련주 부담, 부양책 기대엔 반등 재료."},
    {"date": "2026-07-09", "title": "펩시코 실적발표 (美 어닝시즌 개막)", "category": "실적", "impact": "LOW", "region": "🇺🇸", "tags": ["펩시코", "실적", "어닝시즌"]},
    {"date": "2026-07-10", "title": "TSMC 6월 매출 발표", "category": "실적", "impact": "MEDIUM", "region": "🇹🇼",
     "tags": ["TSMC", "매출", "반도체", "AI"],
     "impact_analysis": "글로벌 반도체 수요 선행지표. 월매출 호조 시 AI·파운드리 수요 확인 → 삼성·SK하이닉스·소부장 동반 강세 신호."},
    {"date": "2026-07-10", "title": "델타항공 실적발표", "category": "실적", "impact": "LOW", "region": "🇺🇸", "tags": ["델타항공", "실적", "여행"]},
    # APEC 정상회담(경주)은 2025년 종료 — 과거 이벤트라 제거(미래 'TBD'로 잘못 잔류했음).
]

# ── 3. FRED 릴리스 캘린더에서 경제지표 발표일 자동 수집 ──────────────
# FRED 경제 릴리스 ID 매핑
# FRED 릴리스 ID (표준): 10=CPI, 46=PPI, 50=고용보고서, 53=GDP, 54=개인소득·PCE, 9=소매판매
# 라벨은 FRED가 반환하는 실제 릴리스명(영문)을 키워드 매칭해 부여 → ID 오류 시 오라벨 방지
FRED_RELEASES = {
    10: {"title": "미국 CPI (소비자물가)",  "category": "경제지표", "impact": "HIGH",   "region": "🇺🇸", "tags": ["CPI", "인플레이션", "Fed"],      "match": "Consumer Price"},
    46: {"title": "미국 PPI (생산자물가)",   "category": "경제지표", "impact": "MEDIUM", "region": "🇺🇸", "tags": ["PPI", "인플레이션"],            "match": "Producer Price"},
    50: {"title": "미국 고용보고서 (NFP)",   "category": "경제지표", "impact": "HIGH",   "region": "🇺🇸", "tags": ["고용", "NFP", "Fed"],            "match": "Employment Situation"},
    53: {"title": "미국 GDP (성장률)",       "category": "경제지표", "impact": "HIGH",   "region": "🇺🇸", "tags": ["GDP", "경기"],                  "match": "Gross Domestic"},
    54: {"title": "미국 PCE·개인소득",       "category": "경제지표", "impact": "HIGH",   "region": "🇺🇸", "tags": ["PCE", "인플레이션", "Fed"],      "match": "Personal Income"},
    9:  {"title": "미국 소매판매",            "category": "경제지표", "impact": "MEDIUM", "region": "🇺🇸", "tags": ["소비", "소매"],                  "match": "Retail"},
}


def _fred_get(url: str):
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=10) as r:
        return json.loads(r.read())


def _release_name(api_key: str, release_id: int):
    """릴리스 ID의 실제 영문명 조회 (ID 검증용)."""
    try:
        params = urllib.parse.urlencode({"release_id": release_id, "file_type": "json", "api_key": api_key})
        data = _fred_get(f"https://api.stlouisfed.org/fred/release?{params}")
        rels = data.get("releases", [])
        return rels[0].get("name") if rels else None
    except Exception:
        return None


def fetch_fred_releases(api_key: str, start: str, end: str) -> list:
    """FRED release/dates API에서 미래 발표 일정을 수집합니다.
    핵심: include_release_dates_with_no_data=true 여야 '미래 예정일'(데이터 없음)이 반환됨."""
    events = []
    for release_id, meta in FRED_RELEASES.items():
        try:
            # ID 검증: 실제 릴리스명이 기대 키워드와 일치하는지
            name = _release_name(api_key, release_id)
            if name and meta.get("match") and meta["match"].lower() not in name.lower():
                print(f"  [WARN] FRED release {release_id} 이름 불일치: '{name}' (기대 '{meta['match']}') → 스킵")
                continue
            params = urllib.parse.urlencode({
                "release_id": release_id,
                "include_release_dates_with_no_data": "true",  # 미래 예정일 포함 (필수)
                "sort_order": "asc",
                "limit": 1000,
                "file_type": "json",
                "api_key": api_key,
            })
            data = _fred_get(f"https://api.stlouisfed.org/fred/release/dates?{params}")
            cnt = 0
            for rd in data.get("release_dates", []):
                d = rd.get("date", "")
                if start <= d <= end:
                    ev = {k: v for k, v in meta.items() if k != "match"}
                    ev.update({
                        "date": d,
                        "time": "08:30 ET",  # BLS/BEA 통상 발표 시각 (카운트다운용)
                        "detail": "BLS/BEA 공식 발표 일정 (FRED 자동수집). 컨센서스 대비 서프라이즈 시 변동성 확대.",
                        "source": "FRED",
                    })
                    events.append(ev)
                    cnt += 1
            print(f"  [FRED] {release_id} {name or meta['title']}: 윈도우 내 {cnt}건")
        except Exception as e:
            print(f"  [WARN] FRED release {release_id} 수집 실패: {e}")
    return events


# ── Nasdaq 경제 캘린더 자동수집 (무료·키 불필요) ──────────────────────
import time as _time

NASDAQ_INDICATORS = [
    {"key": "NFP",    "names": ["nonfarm payrolls", "non farm payrolls"],
     "title": "미국 고용보고서 (NFP)", "impact": "HIGH", "tags": ["고용", "NFP", "Fed"],
     "analysis": "💼 고용 호조(예상 상회) 시 Fed 인하 지연·금리 상승·달러 강세 / 둔화 시 인하 기대·위험자산 선호. 매월 첫 금요일 BLS 발표 — Fed 이중책무의 한 축."},
    {"key": "CPI",    "names": ["cpi", "consumer price index", "inflation rate"],
     "title": "미국 CPI (소비자물가)", "impact": "HIGH", "tags": ["CPI", "인플레이션", "Fed"],
     "analysis": "📊 예상 상회(HOT) 시 달러 강세·미국채 금리 급등·KOSPI 외국인 이탈 / 하회(COOL) 시 인하 기대 강화·나스닥·반도체 수혜. 금리 경로 결정 핵심."},
    {"key": "PPI",    "names": ["ppi", "producer price index"],
     "title": "미국 PPI (생산자물가)", "impact": "MEDIUM", "tags": ["PPI", "인플레이션"],
     "analysis": "🏭 생산자물가는 CPI를 선행. 상승 지속 시 인플레 압력 신호 — 기업 비용 전가 여부 확인."},
    {"key": "GDP",    "names": ["gdp growth rate", "gdp growth rate qoq", "gdp"],
     "title": "미국 GDP (성장률)", "impact": "HIGH", "tags": ["GDP", "경기"],
     "analysis": "📈 성장률 호조 시 경기 견조·위험선호 / 둔화 시 침체 우려·인하 기대. 분기별(속보·잠정·확정) 발표."},
    {"key": "PCE",    "names": ["pce price index", "core pce price index", "core pce price index annual"],
     "title": "미국 PCE 물가", "impact": "HIGH", "tags": ["PCE", "인플레이션", "Fed"],
     "analysis": "🎯 Fed가 가장 중시하는 물가지표. 2% 목표 대비 경로 확인 — 금리 결정에 직결."},
    {"key": "Retail", "names": ["retail sales", "retail sales mom", "retail sales ex autos"],
     "title": "미국 소매판매", "impact": "MEDIUM", "tags": ["소비", "소매"],
     "analysis": "🛒 미국 소비(GDP의 ~70%) 건전성 척도. 강하면 경기 견조·인하 지연 압력."},
    {"key": "FOMC",   "names": ["fed interest rate decision", "federal funds rate", "fed press conference"],
     "title": "미국 FOMC 금리결정", "impact": "HIGH", "tags": ["Fed", "FOMC", "금리"],
     "analysis": "🏦 연준 기준금리 결정·점도표·기자회견. 시장 최대 이벤트 — 금리 경로와 위험자산 방향 결정."},
    {"key": "UMICH",  "names": ["michigan consumer sentiment"],
     "title": "미국 소비자심리(미시간)", "impact": "MEDIUM", "tags": ["심리", "소비심리", "인플레기대"],
     "analysis": "🛍️ 소비심리 + 기대인플레이션. 기대인플레 상승 시 Fed 매파 우려·금리↑. 미시간大 예비치."},
    {"key": "AUC3",   "names": ["3-year note auction"], "category": "국채입찰",
     "title": "미국 3년물 국채입찰", "impact": "LOW", "tags": ["국채", "금리", "입찰"],
     "analysis": "🏦 단기물 수요 확인. 응찰률(bid-to-cover) 약하면 단기금리↑·위험자산 부담. Fed 정책기대 반영."},
    {"key": "AUC10",  "names": ["10-year note auction"], "category": "국채입찰",
     "title": "미국 10년물 국채입찰", "impact": "MEDIUM", "tags": ["국채", "금리", "입찰"],
     "analysis": "🏦 벤치마크 10년물 수요. 약한 입찰 → 장기금리↑ → 성장주·나스닥 부담."},
    {"key": "AUC30",  "names": ["30-year bond auction"], "category": "국채입찰",
     "title": "미국 30년물 국채입찰", "impact": "MEDIUM", "tags": ["국채", "금리", "입찰"],
     "analysis": "🏦 초장기물. 약한 입찰 → 장기금리 급등 → 성장주 직격·재정적자 우려(bond vigilante)."},
]


def _clean_cell(s) -> str:
    return (s or "").replace("&nbsp;", "").replace("\xa0", "").strip()


def fetch_nasdaq_calendar(days_back: int = 10, days_ahead: int = 50) -> list:
    """Nasdaq 경제 캘린더 API(무료·키 불필요)에서 미국 주요 지표 발표 일정을 수집.
    공식 일정 + 컨센서스 + 직전값 + (발표 시) 실제값 제공."""
    events = []
    base = date.today()
    ok_days = 0
    for off in range(-days_back, days_ahead + 1):
        d = base + timedelta(days=off)
        if d.weekday() >= 5:  # 주말 스킵 (미 지표는 평일)
            continue
        ds = d.isoformat()
        try:
            url = f"https://api.nasdaq.com/api/calendar/economicevents?date={ds}"
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (compatible; ai-finance/1.0)",
                "Accept": "application/json",
            })
            with urllib.request.urlopen(req, timeout=12) as r:
                payload = json.loads(r.read())
            rows = (payload.get("data") or {}).get("rows") or []
            ok_days += 1
        except Exception:
            continue
        us = [row for row in rows if row.get("country") == "United States"]
        for ind in NASDAQ_INDICATORS:
            match = None
            for row in us:
                if (row.get("eventName") or "").strip().lower() in ind["names"]:
                    match = row
                    break
            if not match:
                continue
            gmt = _clean_cell(match.get("gmt"))
            ev = {
                "date": ds,
                "time": f"{gmt} ET" if gmt else "08:30 ET",
                "title": ind["title"], "category": ind.get("category", "경제지표"), "impact": ind["impact"],
                "region": "🇺🇸", "tags": ind["tags"][:],
                "impact_analysis": ind["analysis"],
                "detail": "Nasdaq 경제캘린더 자동수집. 컨센서스 대비 서프라이즈 시 변동성 확대.",
                "source": "Nasdaq",
            }
            cons = _clean_cell(match.get("consensus"))
            prev = _clean_cell(match.get("previous"))
            act = _clean_cell(match.get("actual"))
            if cons:
                ev["consensus"] = cons
            if prev:
                ev["prior_hint"] = prev
            if act:
                ev["nasdaq_actual"] = act
            events.append(ev)
        _time.sleep(0.12)
    print(f"  [Nasdaq] {ok_days}일 조회, {len(events)}개 미국 지표 이벤트 수집")
    return events


# ── 주요 실적 발표 (Nasdaq earnings, AI·반도체 워치리스트) ─────────────
EARNINGS_WATCH = {
    "ORCL": ("오라클",   "HIGH",   "☁️ AI 클라우드(OCI)·RPO(잔여수주) 가이던스 = AI 데이터센터 수요 바로미터. 강하면 엔비디아·SK하이닉스 HBM 동반 강세."),
    "TSM":  ("TSMC",     "HIGH",   "🏭 파운드리 1위. AI칩 수요·가이던스 → SOX·삼성전자·SK하이닉스 직결."),
    "NVDA": ("엔비디아",  "HIGH",   "🚀 AI 가속기 대장. 데이터센터 매출·가이던스가 전체 AI 테마 좌우."),
    "AVGO": ("브로드컴",  "HIGH",   "🔌 커스텀 AI칩·네트워킹. 하이퍼스케일러 수요 지표."),
    "MU":   ("마이크론",  "MEDIUM", "💾 HBM·D램. AI 메모리 사이클 — 삼성·SK하이닉스 동조."),
    "AMD":  ("AMD",      "MEDIUM", "⚙️ MI 가속기·CPU. 엔비디아 대항 점유율."),
    "ADBE": ("어도비",    "MEDIUM", "🎨 크리에이티브 SaaS·AI(Firefly) 수익화."),
}


def fetch_nasdaq_earnings(start: str, end: str) -> list:
    """Nasdaq earnings 캘린더에서 워치리스트 종목 실적일 수집 (무료·키 불필요)."""
    events = []
    d = date.fromisoformat(start)
    end_d = date.fromisoformat(end)
    cnt = 0
    while d <= end_d:
        if d.weekday() < 5:
            try:
                url = f"https://api.nasdaq.com/api/calendar/earnings?date={d.isoformat()}"
                req = urllib.request.Request(url, headers={
                    "User-Agent": "Mozilla/5.0 (compatible; ai-finance/1.0)", "Accept": "application/json"})
                with urllib.request.urlopen(req, timeout=12) as r:
                    rows = (json.loads(r.read()).get("data") or {}).get("rows") or []
                for row in rows:
                    sym = row.get("symbol")
                    if sym in EARNINGS_WATCH:
                        name, impact, analysis = EARNINGS_WATCH[sym]
                        when = (row.get("time") or "")
                        tlabel = "장마감 후" if "after" in when else "장전" if "before" in when else ""
                        eps = _clean_cell(row.get("epsForecast"))
                        events.append({
                            "date": d.isoformat(), "time": "",
                            "title": f"{name} 실적발표", "category": "실적", "impact": impact,
                            "region": "🇺🇸", "tags": ["실적", sym, "AI"],
                            "impact_analysis": analysis,
                            "detail": f"{sym} 분기 실적{(' (' + tlabel + ')') if tlabel else ''}." + (f" EPS 컨센 {eps}." if eps else ""),
                            "consensus": (f"EPS {eps}" if eps else None),
                            "source": "Nasdaq",
                        })
                        cnt += 1
                _time.sleep(0.12)
            except Exception:
                pass
        d += timedelta(days=1)
    print(f"  [실적] 워치리스트 {cnt}건 수집")
    return events


# ── 한국 파생 만기·지수 정기변경 (규칙 기반 자동생성) ─────────────────
def korean_derivative_events(start: str, end: str) -> list:
    """분기(3·6·9·12월) 둘째 목요일 = 동시만기 + (6·12월) 지수 정기변경."""
    events = []
    today = date.today()
    for year in (today.year, today.year + 1):
        for month in (3, 6, 9, 12):
            first = date(year, month, 1)
            offset = (3 - first.weekday()) % 7  # 목요일=3
            second_thu = first + timedelta(days=offset + 7)
            ds = second_thu.isoformat()
            if not (start <= ds <= end):
                continue
            events.append({
                "date": ds, "time": "15:20 KST",
                "title": "한국 선물·옵션 동시만기 (네 마녀의 날)", "category": "파생만기", "impact": "HIGH",
                "region": "🇰🇷", "tags": ["만기", "파생", "수급", "변동성"],
                "impact_analysis": "🎭 분기 동시만기 → 프로그램 매물·수급 왜곡·변동성 급증. 외국인 선물청산 시 현물 충격. 만기 당일 장마감 동시호가(15:20~) 주의.",
                "detail": "주가지수·개별주식 선물/옵션 동시 만기 (분기 둘째 목요일).",
                "source": "규칙생성",
            })
            if month in (6, 12):
                events.append({
                    "date": ds, "time": "",
                    "title": "코스피200·코스닥150 정기변경", "category": "파생만기", "impact": "MEDIUM",
                    "region": "🇰🇷", "tags": ["리밸런싱", "패시브", "수급"],
                    "impact_analysis": "📊 지수 편입/편출 → 패시브 펀드 기계적 매매(편입 매수·편출 매도). 동시만기일과 겹쳐 수급 변동 증폭.",
                    "detail": "한국거래소 반기 정기변경 적용일(둘째 목요일).",
                    "source": "규칙생성",
                })
    return events


# ── TSMC 월간 매출 (규칙 기반 — 매월 ~10일 전월 매출 공시) ─────────────
def tsmc_revenue_events(start: str, end: str) -> list:
    """TSMC는 매월 10일경 전월 매출을 공시 — 반도체/AI 수요 선행지표."""
    events = []
    today = date.today()
    base_idx = (today.year * 12) + (today.month - 1)  # 0-indexed 월
    for off in range(-1, 4):  # 지난달 ~ 향후 3개월
        idx = base_idx + off
        y, mm = idx // 12, idx % 12 + 1
        ds = date(y, mm, 10).isoformat()
        if not (start <= ds <= end):
            continue
        prev = mm - 1 if mm > 1 else 12
        events.append({
            "date": ds, "time": "",
            "title": f"TSMC {prev}월 매출 발표", "category": "실적", "impact": "HIGH",
            "region": "🇹🇼", "tags": ["TSMC", "반도체", "AI", "매출"],
            "impact_analysis": "🏭 TSMC 월간 매출 = 반도체 수요 선행지표. AI칩(엔비디아·애플) 주문 강도 확인 → SOX·삼성전자·SK하이닉스 직결. YoY·MoM 급증 시 AI 슈퍼사이클 신호.",
            "detail": "대만 TSMC 전월 매출 공시 (통상 매월 10일경). AI 파운드리 수요 풍향계.",
            "source": "규칙생성",
        })
    return events


# ── 발표 완료 실제값 수집 (FRED 시리즈) ───────────────────────────────
# 태그 키워드 → (series_id, metric, 단위라벨, 소수자리)
#   metric: yoy=전년동월비%, mom=전월대비변화, level=원시값
#   (units= 파라미터는 PAYEMS에서 400 오류 → 원시값 받아 직접 계산)
FRED_VALUE_MAP = [
    (["CPI"],        "CPIAUCSL", "yoy",  "% YoY", 1),  # 소비자물가 전년비
    (["PCE"],        "PCEPI",    "yoy",  "% YoY", 1),  # PCE 물가 전년비
    (["PPI"],        "PPIFIS",   "yoy",  "% YoY", 1),  # 생산자물가 전년비
    (["고용", "NFP"], "PAYEMS",   "mom",  "천명",  0),  # 비농업고용 전월대비(천명)
    (["GDP"],        "A191RL1Q225SBEA", "level", "%", 1),  # 실질GDP 성장률(연율)
    (["소매", "소비"], "RSAFS",    "yoy",  "% YoY", 1),  # 소매판매 전년비
]


def fetch_indicator_value(api_key: str, series_id: str, metric: str) -> tuple:
    """FRED 원시 관측값을 받아 metric(yoy/mom/level) 계산.
    (latest_val, latest_period_date, prior_val) 또는 (None,None,None).
    units= 파라미터 미사용(400 회피) — m2_monitor와 동일한 안정적 호출.
    """
    try:
        from datetime import date as _date
        start = f"{_date.today().year - 2}-01-01"
        params = urllib.parse.urlencode({
            "series_id": series_id, "api_key": api_key.strip(),
            "observation_start": start, "file_type": "json", "sort_order": "asc",
        })
        url = f"https://api.stlouisfed.org/fred/series/observations?{params}"
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=15) as r:
            data = json.loads(r.read())
        rows = [(o["date"], float(o["value"])) for o in data.get("observations", [])
                if o.get("value") not in (".", "", None)]
        if len(rows) < 2:
            return None, None, None

        if metric == "level":
            latest_d, latest_v = rows[-1]
            prior_v = rows[-2][1]
            return latest_v, latest_d, prior_v

        if metric == "mom":
            latest_d, latest_v = rows[-1]
            prior_v = rows[-2][1]
            # 전월대비 변화 (천명 등): 절대 변화량
            cur_chg = latest_v - prior_v
            prv_chg = prior_v - rows[-3][1] if len(rows) >= 3 else None
            return cur_chg, latest_d, prv_chg

        # yoy: 전년동월비 %
        date_to_val = {d: v for d, v in rows}
        def _yoy(idx):
            d, v = rows[idx]
            ya = f"{int(d[:4]) - 1}{d[4:]}"
            pv = date_to_val.get(ya)
            return ((v - pv) / pv * 100) if (pv and pv != 0) else None
        cur = _yoy(-1)
        prv = _yoy(-2)
        if cur is None:
            return None, None, None
        return cur, rows[-1][0], prv
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

    # 미국 경제지표·국채입찰 (Nasdaq 캘린더 자동 수집 — 무료·키 불필요)
    print("[Nasdaq] 미국 경제지표·국채입찰 자동 수집 중...")
    nasdaq_events = fetch_nasdaq_calendar(days_back=10, days_ahead=50)
    nasdaq_win = filter_window(nasdaq_events)
    all_events.extend(nasdaq_win)
    print(f"[Nasdaq] {len(nasdaq_win)}개 경제지표·입찰 (윈도우 내)")

    # 주요 실적 (Nasdaq earnings — AI·반도체 워치리스트, 향후 28일)
    print("[실적] 주요 종목 실적일 수집 중...")
    earn_start = today.strftime("%Y-%m-%d")
    earn_end = (today + timedelta(days=28)).strftime("%Y-%m-%d")
    earnings_events = fetch_nasdaq_earnings(earn_start, earn_end)
    all_events.extend(filter_window(earnings_events))

    # 한국 파생 만기·지수 정기변경 (규칙 생성)
    past_str = (today - timedelta(days=14)).strftime("%Y-%m-%d")
    kr_deriv = filter_window(korean_derivative_events(past_str, end_str))
    all_events.extend(kr_deriv)
    print(f"[한국 파생] {len(kr_deriv)}개 만기·리밸런싱 (윈도우 내)")

    # TSMC 월간 매출 (규칙 생성)
    tsmc = filter_window(tsmc_revenue_events(past_str, end_str))
    all_events.extend(tsmc)
    print(f"[TSMC] {len(tsmc)}개 월매출 일정 (윈도우 내)")

    # 수동 큐레이션 (IPO/지정학)
    curated = filter_upcoming(CURATED_EVENTS)
    all_events.extend(curated)
    print(f"[큐레이션] {len(curated)}개 IPO/지정학 이벤트 로드")

    # 중복 제거 — 경제지표는 (날짜+지표)로 묶어 FRED·큐레이션 병합 (분석 풍부한 쪽 우선)
    _INDICATORS = ["CPI", "PPI", "NFP", "고용", "GDP", "PCE", "소매", "소비"]

    def _dedup_key(ev):
        d = ev.get("date")
        if ev.get("category") == "경제지표":
            tags = ev.get("tags", [])
            ind = next((i for i in _INDICATORS if i in tags), None)
            if ind in ("고용", "NFP"):
                ind = "NFP"
            if ind in ("소매", "소비"):
                ind = "소매"
            if ind:
                return (d, "경제지표", ind)
        return (d, ev.get("title"))

    seen: dict = {}  # key → index in deduped
    deduped = []
    for ev in all_events:
        key = _dedup_key(ev)
        if key not in seen:
            seen[key] = len(deduped)
            deduped.append(ev)
        else:
            # 더 풍부한 항목(impact_analysis 등) 우선, 단 FRED date·source는 보존
            existing_idx = seen[key]
            cur = deduped[existing_idx]
            richer = ev if len(ev) > len(cur) else cur
            poorer = cur if richer is ev else ev
            # FRED source 정보가 한쪽에만 있으면 병합
            if poorer.get("source") == "FRED" and "source" not in richer:
                richer = {**richer, "source": "FRED"}
            deduped[existing_idx] = richer

    # 발표완료 경제지표에 실제값 부여 (FRED)
    print("\n[발표완료 결과 수집]")
    deduped = enrich_events_with_actuals(deduped, api_key)

    # Nasdaq actual 폴백: 발표완료인데 FRED 실제값이 없으면 Nasdaq 값 사용
    for ev in deduped:
        if ev.get("status") == "released" and ev.get("actual_value") is None and ev.get("nasdaq_actual"):
            ev["actual_value"] = ev["nasdaq_actual"]
            ev["actual_unit"] = ""

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
