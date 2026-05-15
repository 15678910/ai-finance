"""
반도체 챌린저 모니터 — 신흥 위협 분석 (Phase 1: 휴리스틱 기반)
=================================================================

목적:
  Cerebras, Groq, Tenstorrent 등 신흥 반도체 기업이 기존 사업자
  (NVIDIA, 삼성전자, SK하이닉스, 한미반도체, TSMC)에 미치는 영향을 추적.

데이터 수집:
  1. 상장 챌린저(Cerebras, AMD 등): yfinance에서 현재가/시총 자동 수집
  2. 비상장 챌린저: 사전 정의된 프로파일 + 자금조달 정보
  3. 실시간 뉴스: 회사명 키워드로 RSS 매칭 → 한국어 자동 번역
  4. 영향 분석: 사람이 사전 정의한 thesis (분기마다 업데이트)

추적 카테고리:
  · AI Accelerator (training/inference)
  · Edge AI / 뉴로모픽
  · 광학 컴퓨팅 (Photonic)
  · 차세대 CPU/RISC-V
  · Transformer-specific ASIC

알림 트리거: 챌린저 관련 신규 뉴스 + 주요 이벤트 (펀딩·IPO·대형 계약)

🚨 시뮬레이션. 자동 매매 금지.
"""

import os
import sys
import json
import re
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta

try:
    from defusedxml import ElementTree as ET  # type: ignore
except ImportError:
    import xml.etree.ElementTree as ET

try:
    import yfinance as yf
except ImportError as e:
    print(f"[오류] yfinance 미설치: {e}")
    sys.exit(1)

from core import send_message, load_state, save_state, translate_to_korean

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "semi_challengers.json")
STATE_NAME = "semi_challengers"
KST = timezone(timedelta(hours=9))

USER_AGENT = "ai-finance-challenger-monitor/1.0"
TIMEOUT_SEC = 15

# ====================================================================
# 챌린저 기업 프로파일 (Phase 1: 수동 정의, 분기마다 업데이트)
# ====================================================================
INCUMBENTS = {
    "NVDA":   {"name": "NVIDIA",       "market": "US",    "role": "AI GPU 절대강자"},
    "AMD":    {"name": "AMD",          "market": "US",    "role": "기존 #2 챌린저 (MI300X)"},
    "INTC":   {"name": "Intel",        "market": "US",    "role": "기존 #3 (Gaudi)"},
    "TSM":    {"name": "TSMC",         "market": "US",    "role": "파운드리 절대강자"},
    "005930": {"name": "삼성전자",      "market": "KOSPI", "role": "메모리·파운드리"},
    "000660": {"name": "SK하이닉스",    "market": "KOSPI", "role": "HBM 시장 1위"},
    "042700": {"name": "한미반도체",    "market": "KOSPI", "role": "HBM TC 본더"},
    "000990": {"name": "DB하이텍",      "market": "KOSPI", "role": "8인치 파운드리"},
    "ASML":   {"name": "ASML",         "market": "US",    "role": "EUV 노광장비"},
}

CHALLENGERS = [
    {
        "id": "cerebras",
        "name": "Cerebras Systems",
        "ticker": "CBRS",  # NASDAQ 2024 IPO
        "category": "AI Accelerator",
        "tech_label": "Wafer-scale Engine (WSE-3)",
        "founded": 2016,
        "headquarters": "Sunnyvale, CA",
        "production_at": "TSMC 5nm",
        "tech_description": "단일 12인치 웨이퍼에 90만 코어 집적, 인터칩 통신 병목 제거. 대규모 LLM 훈련에 특화.",
        "key_advantage": "온-웨이퍼 SRAM 44GB → HBM 비의존 + 통신 지연 zero",
        "funding_total_usd_m": 720,
        "valuation_usd_b": 4.0,
        "status": "Public (2024 IPO)",
        "key_customers": ["G42 (UAE)", "Mayo Clinic", "Lawrence Livermore"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "moderate", "reasoning": "대규모 LLM 훈련 시장 H100 직접 경쟁자. CUDA 생태계 lock-in으로 단기 침투 제한적이나 G42·Mayo 같은 대형 계약으로 점진적 잠식."},
            "AMD":    {"sentiment": "negative", "magnitude": "small",    "reasoning": "AMD도 동일 시장 노리는 #2 챌린저 — Cerebras에 일부 잠식 가능."},
            "TSM":    {"sentiment": "positive", "magnitude": "small",    "reasoning": "Cerebras도 TSMC에서 위탁 생산 → 점유율 확대 시 수주 증가."},
            "005930": {"sentiment": "neutral",  "magnitude": "small",    "reasoning": "WSE-3는 온-웨이퍼 SRAM 사용 → 단기 HBM 영향 제한. 장기적으로 비-HBM 아키텍처 확산 시 부정적."},
            "000660": {"sentiment": "negative", "magnitude": "moderate", "reasoning": "HBM 매출 비중 50%+ → Cerebras 등 비-HBM 아키텍처 점유율 확대 시 직접 위협."},
            "042700": {"sentiment": "negative", "magnitude": "large",    "reasoning": "HBM TC 본더 매출 핵심 → 비-HBM 칩 확산은 가장 큰 타격 종목."},
        },
        "rss_keywords": ["Cerebras", "WSE-3", "wafer-scale"],
        "tags": ["AI", "training", "wafer-scale", "비-HBM"],
    },
    {
        "id": "groq",
        "name": "Groq",
        "ticker": None,  # 비상장
        "category": "AI Inference Accelerator",
        "tech_label": "LPU (Language Processing Unit)",
        "founded": 2016,
        "headquarters": "Mountain View, CA",
        "production_at": "GlobalFoundries 14nm (구세대), Samsung 4nm (차세대)",
        "tech_description": "LLM 추론 특화 ASIC. 결정론적 실행 + 외부 메모리 없는 SRAM 온칩 설계로 토큰 생성 속도 10x.",
        "key_advantage": "추론 latency 시장에서 NVIDIA 대비 10-20x 빠른 토큰 생성",
        "funding_total_usd_m": 1640,
        "valuation_usd_b": 6.9,
        "status": "Private (BlackRock 2024 라운드)",
        "key_customers": ["Aramco", "Meta(부분)", "Yann LeCun 추천"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "large",    "reasoning": "추론(inference) 시장이 훈련보다 큰 매출 (NVIDIA 매출의 50%+) → Groq 점유율 확대 시 직접 위협."},
            "AMD":    {"sentiment": "negative", "magnitude": "small",    "reasoning": "AMD MI300X도 추론 시장 노림. Groq에 일부 점유율 빼앗길 수 있음."},
            "INTC":   {"sentiment": "negative", "magnitude": "small",    "reasoning": "Intel Gaudi의 추론 시장 진입 어렵게 됨."},
            "TSM":    {"sentiment": "neutral",  "magnitude": "small",    "reasoning": "Groq는 Samsung 파운드리 이전 → TSMC에는 약한 부정."},
            "005930": {"sentiment": "positive", "magnitude": "moderate", "reasoning": "Groq 차세대 칩 Samsung Foundry 4nm 위탁 → 파운드리 수주 증가."},
            "000660": {"sentiment": "negative", "magnitude": "moderate", "reasoning": "Groq LPU는 SRAM 온칩 → HBM 미사용. 추론 시장에서 HBM 비중 감소 위협."},
            "042700": {"sentiment": "negative", "magnitude": "moderate", "reasoning": "추론 시장 비-HBM 비중 증가 → 한미 HBM 본더 수주 위축 우려."},
        },
        "rss_keywords": ["Groq", "Groq LPU", "Language Processing Unit"],
        "tags": ["AI", "inference", "ASIC", "비-HBM"],
    },
    {
        "id": "sambanova",
        "name": "SambaNova Systems",
        "ticker": None,
        "category": "AI Dataflow Architecture",
        "tech_label": "RDU (Reconfigurable Dataflow Unit)",
        "founded": 2017,
        "headquarters": "Palo Alto, CA",
        "production_at": "TSMC",
        "tech_description": "데이터플로 아키텍처로 LLM 훈련+추론 통합. 엔터프라이즈 RAG 솔루션에 특화.",
        "key_advantage": "데이터플로 + HBM3 + 1TB DDR → 1조 파라미터 LLM 추론 가능",
        "funding_total_usd_m": 1100,
        "valuation_usd_b": 5.1,
        "status": "Private",
        "key_customers": ["Argonne", "LLNL", "Saudi Aramco"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "moderate", "reasoning": "엔터프라이즈 LLM 시장 직접 경쟁. NVIDIA DGX 시스템 대안."},
            "TSM":    {"sentiment": "positive", "magnitude": "small",    "reasoning": "TSMC 생산 → 점유율 확대 시 수주 증가."},
            "000660": {"sentiment": "positive", "magnitude": "small",    "reasoning": "SambaNova는 HBM3 사용 → HBM 수요 유지."},
            "042700": {"sentiment": "positive", "magnitude": "small",    "reasoning": "HBM 본더 수요 유지."},
        },
        "rss_keywords": ["SambaNova", "RDU"],
        "tags": ["AI", "training", "enterprise", "HBM 사용"],
    },
    {
        "id": "tenstorrent",
        "name": "Tenstorrent",
        "ticker": None,
        "category": "AI + RISC-V",
        "tech_label": "Black hole + Wormhole",
        "founded": 2016,
        "headquarters": "Toronto, Canada",
        "production_at": "TSMC + Samsung Foundry",
        "tech_description": "Jim Keller(전 Apple/Tesla/AMD) 주도. RISC-V CPU + AI 가속기 통합. 오픈소스 SW 스택.",
        "key_advantage": "오픈 아키텍처 (CUDA 대안) + 합리적 가격",
        "funding_total_usd_m": 850,
        "valuation_usd_b": 2.6,
        "status": "Private (2024 Samsung·Hyundai 투자)",
        "key_customers": ["Hyundai", "LG (개발 협업)"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "moderate", "reasoning": "Jim Keller 명성 + 오픈 SW 스택으로 CUDA 의존도 줄이려는 고객 유치."},
            "ARM":    {"sentiment": "negative", "magnitude": "moderate", "reasoning": "RISC-V 진영의 대표주자 — ARM 라이센스 모델에 위협."},
            "TSM":    {"sentiment": "positive", "magnitude": "small",    "reasoning": "TSMC 위탁 → 점유율 확대 시 수주 증가."},
            "005930": {"sentiment": "positive", "magnitude": "moderate", "reasoning": "Samsung Foundry에서 일부 생산 + Samsung 전략 투자자 → 한국 수혜."},
            "000660": {"sentiment": "neutral",  "magnitude": "small",    "reasoning": "HBM 사용 일부 — 영향 제한적."},
        },
        "rss_keywords": ["Tenstorrent", "Jim Keller", "Black hole AI"],
        "tags": ["AI", "RISC-V", "오픈SW", "Korean tie"],
    },
    {
        "id": "lightmatter",
        "name": "Lightmatter",
        "ticker": None,
        "category": "Photonic Computing",
        "tech_label": "Envise + Passage (광학 칩)",
        "founded": 2017,
        "headquarters": "Boston, MA",
        "production_at": "GlobalFoundries",
        "tech_description": "실리콘 포토닉스 기반 광학 AI 연산. 전기 신호 대신 빛으로 행렬곱 → 100x 에너지 효율.",
        "key_advantage": "데이터센터 전력 소비 폭발 문제의 잠재적 해법",
        "funding_total_usd_m": 270,
        "valuation_usd_b": 1.2,
        "status": "Private",
        "key_customers": ["TBD (PoC 단계)"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "small",    "reasoning": "장기 위협. 5-10년 후 광학 컴퓨팅 상용화 시 GPU 시장 재편 가능. 단기 영향 제한."},
            "005930": {"sentiment": "negative", "magnitude": "small",    "reasoning": "전통 메모리 패러다임 변화 위협 (장기). 광학 컴퓨팅은 메모리 의존도 낮음."},
            "000660": {"sentiment": "negative", "magnitude": "small",    "reasoning": "장기 HBM 수요 위협."},
            "042700": {"sentiment": "neutral",  "magnitude": "small",    "reasoning": "단기 영향 없음."},
        },
        "rss_keywords": ["Lightmatter", "photonic AI", "silicon photonics"],
        "tags": ["AI", "photonic", "장기위협"],
    },
    {
        "id": "etched",
        "name": "Etched.ai",
        "ticker": None,
        "category": "Transformer-Specific ASIC",
        "tech_label": "Sohu",
        "founded": 2022,
        "headquarters": "San Francisco, CA",
        "production_at": "TSMC",
        "tech_description": "Transformer 아키텍처 hardwired ASIC. LLM 외 다른 모델 지원 불가. 추론 비용 10x 절감.",
        "key_advantage": "Transformer만 지원 → 일반 GPU 대비 10x 효율",
        "funding_total_usd_m": 120,
        "valuation_usd_b": 0.5,
        "status": "Private",
        "key_customers": ["TBD"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "moderate", "reasoning": "Transformer 추론 시장의 잠재적 disruptor — Sohu 양산 성공 시 큰 위협."},
            "TSM":    {"sentiment": "positive", "magnitude": "small",    "reasoning": "TSMC 위탁."},
            "005930": {"sentiment": "neutral",  "magnitude": "small",    "reasoning": "HBM 사용 → 메모리 수요 유지."},
            "000660": {"sentiment": "neutral",  "magnitude": "small",    "reasoning": "HBM 사용."},
        },
        "rss_keywords": ["Etched.ai", "Etched Sohu", "Sohu chip", "Transformer ASIC"],
        "tags": ["AI", "inference", "ASIC", "Transformer전용"],
    },
    {
        "id": "rain_ai",
        "name": "Rain AI",
        "ticker": None,
        "category": "Neuromorphic",
        "tech_label": "RNNs + Memristor",
        "founded": 2017,
        "headquarters": "San Francisco, CA",
        "production_at": "(R&D)",
        "tech_description": "뉴로모픽 컴퓨팅 — 뇌 시냅스 모방. OpenAI Sam Altman 개인 투자.",
        "key_advantage": "에너지 효율 100x (목표) + 엣지 AI",
        "funding_total_usd_m": 50,
        "valuation_usd_b": 0.15,
        "status": "Private (Pre-revenue)",
        "key_customers": [],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "small",    "reasoning": "장기 (5-10년) 엣지 AI 시장 위협."},
            "INTC":   {"sentiment": "negative", "magnitude": "small",    "reasoning": "Intel Loihi 뉴로모픽도 경쟁."},
        },
        "rss_keywords": ["Rain AI", "Rain Neuromorphics", "neuromorphic chip Altman"],
        "tags": ["AI", "neuromorphic", "장기"],
    },
    {
        "id": "d_matrix",
        "name": "d-Matrix",
        "ticker": None,
        "category": "LLM Inference",
        "tech_label": "Corsair (3D-stacked DIMC)",
        "founded": 2019,
        "headquarters": "Santa Clara, CA",
        "production_at": "TSMC",
        "tech_description": "Digital In-Memory Computing (DIMC) — 메모리 안에서 직접 연산. 추론 시 데이터 이동 최소화.",
        "key_advantage": "추론 에너지 효율 10x + DDR5/HBM 인터페이스 유지",
        "funding_total_usd_m": 160,
        "valuation_usd_b": 1.0,
        "status": "Private (Microsoft 투자)",
        "key_customers": ["Microsoft Azure (테스트)"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "moderate", "reasoning": "추론 시장 직접 경쟁. Microsoft Azure 도입 검토 = 큰 잠재 매출."},
            "TSM":    {"sentiment": "positive", "magnitude": "small",    "reasoning": "TSMC 생산."},
            "000660": {"sentiment": "neutral",  "magnitude": "small",    "reasoning": "HBM 인터페이스 유지 → 영향 제한."},
            "042700": {"sentiment": "neutral",  "magnitude": "small",    "reasoning": "HBM 본더 일부 수요 유지."},
        },
        "rss_keywords": ["d-Matrix", "dMatrix Corsair", "DIMC chip"],
        "tags": ["AI", "inference", "DIMC", "Microsoft"],
    },
    {
        "id": "matx",
        "name": "MatX",
        "ticker": None,
        "category": "LLM-Specific ASIC",
        "tech_label": "LLM-optimized accelerator",
        "founded": 2022,
        "headquarters": "Mountain View, CA",
        "production_at": "TSMC",
        "tech_description": "전 Google·OpenAI 출신 창업. LLM 훈련+추론 통합 ASIC. 비공개 단계.",
        "key_advantage": "LLM 전용 설계 → 효율 극대화 (예상)",
        "funding_total_usd_m": 25,
        "valuation_usd_b": 0.3,
        "status": "Private (Stealth)",
        "key_customers": [],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "small",    "reasoning": "초기 단계 — 영향 제한적. 단 출신 배경 + Hottest 자금조달 라운드로 주목."},
            "TSM":    {"sentiment": "positive", "magnitude": "small",    "reasoning": "TSMC 위탁 추정."},
        },
        "rss_keywords": ["MatX chip", "MatX startup", "LLM ASIC startup"],
        "tags": ["AI", "ASIC", "stealth"],
    },
    {
        "id": "graphcore",
        "name": "Graphcore",
        "ticker": None,
        "category": "AI Accelerator (IPU)",
        "tech_label": "Bow IPU + Mk3 (개발 중)",
        "founded": 2016,
        "headquarters": "Bristol, UK",
        "production_at": "TSMC",
        "tech_description": "Intelligence Processing Unit. 그래프 기반 연산 아키텍처. Softbank 2024년 인수.",
        "key_advantage": "MIMD 아키텍처 — sparse 모델에 강점",
        "funding_total_usd_m": 700,
        "valuation_usd_b": 0.5,  # 인수 가격 하락
        "status": "Acquired by Softbank (2024)",
        "key_customers": ["일부 유럽 학술기관"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "small",    "reasoning": "한때 NVIDIA 대안으로 주목 받았으나 점유율 미미. Softbank 인수 후 재도약 가능성 있으나 단기 영향 적음."},
        },
        "rss_keywords": ["Graphcore", "Graphcore IPU", "Softbank Graphcore"],
        "tags": ["AI", "IPU", "Softbank"],
    },
    {
        "id": "amd_mi300",
        "name": "AMD MI300X / MI325X",
        "ticker": "AMD",
        "category": "AI GPU (현직 #2)",
        "tech_label": "Instinct MI300X / MI325X",
        "founded": 1969,
        "headquarters": "Santa Clara, CA",
        "production_at": "TSMC",
        "tech_description": "기존 챌린저 — NVIDIA H100 직접 경쟁. CDNA3/4 아키텍처 + 192GB HBM3.",
        "key_advantage": "메모리 용량 우위 (192GB vs H100 80GB) → 대형 모델 추론 유리",
        "status": "Public ($AMD)",
        "key_customers": ["Microsoft", "Meta", "Oracle", "OpenAI(테스트)"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "moderate", "reasoning": "데이터센터 GPU 시장 유일한 의미 있는 #2. MI300X 매출 빠른 성장 ($5B+ 2024)."},
            "TSM":    {"sentiment": "positive", "magnitude": "moderate", "reasoning": "AMD 점유율 확대 = TSMC 수주 증가."},
            "005930": {"sentiment": "positive", "magnitude": "moderate", "reasoning": "MI300X HBM3 사용 → Samsung HBM 수요 증가."},
            "000660": {"sentiment": "positive", "magnitude": "large",    "reasoning": "MI325X 12-Hi HBM3E 채택 → SK하이닉스 핵심 수혜자."},
            "042700": {"sentiment": "positive", "magnitude": "moderate", "reasoning": "HBM 본더 수요 증가."},
        },
        "rss_keywords": ["AMD MI300", "AMD MI325", "AMD Instinct"],
        "tags": ["AI", "GPU", "HBM 수혜", "현직"],
    },
    {
        "id": "mythic",
        "name": "Mythic AI",
        "ticker": None,
        "category": "Edge AI (Analog)",
        "tech_label": "M1076 (아날로그 행렬곱)",
        "founded": 2012,
        "headquarters": "Redwood City, CA",
        "production_at": "GlobalFoundries 40nm",
        "tech_description": "아날로그 in-memory 컴퓨팅. 엣지 AI 추론 1W 미만. 2022년 자금난, 2024년 재기.",
        "key_advantage": "초저전력 (전통 GPU의 100x 효율) — 엣지/IoT 특화",
        "funding_total_usd_m": 165,
        "valuation_usd_b": 0.4,
        "status": "Private (2024 재투자)",
        "key_customers": ["보안카메라 OEM 일부"],
        "impact_thesis": {
            "NVDA":   {"sentiment": "negative", "magnitude": "small",    "reasoning": "NVIDIA Jetson 엣지 시장 잠재 위협. 데이터센터 영향 없음."},
        },
        "rss_keywords": ["Mythic AI", "Mythic M1076", "analog AI chip"],
        "tags": ["edge AI", "analog", "low power"],
    },
]

# ====================================================================
# RSS 피드 (반도체·AI 기술 분야 집중)
# ====================================================================
RSS_FEEDS = [
    {"name": "CoinDesk",             "url": "https://www.coindesk.com/arc/outboundfeeds/rss/",         "lang": "en"},
    {"name": "Decrypt",              "url": "https://decrypt.co/feed",                                 "lang": "en"},
    {"name": "The Block",            "url": "https://www.theblock.co/rss.xml",                         "lang": "en"},
    {"name": "Heisenberg(기술분석)",  "url": "https://heisenberg.kr/feed/",                             "lang": "ko"},
    # Google News 검색 (반도체 startup 키워드)
    {"name": "Google News(AI반도체)", "url": "https://news.google.com/rss/search?q=AI+chip+startup+OR+Cerebras+OR+Groq+OR+Tenstorrent&hl=en-US&gl=US&ceid=US:en", "lang": "en"},
    {"name": "Google News(한국)",    "url": "https://news.google.com/rss/search?q=Cerebras+OR+Groq+OR+%EB%B0%98%EB%8F%84%EC%B2%B4+%EC%8A%A4%ED%83%80%ED%8A%B8%EC%97%85+when%3A7d&hl=ko&gl=KR&ceid=KR:ko", "lang": "ko"},
]


# ====================================================================
# HTTP 헬퍼
# ====================================================================
def http_get(url: str, timeout: int = TIMEOUT_SEC) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:  # nosec
        return resp.read()


def parse_rss_items(xml_bytes: bytes) -> list:
    """RSS/Atom → [{title, link, summary, pub_date, guid}]"""
    items = []
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return items
    # RSS 2.0
    for item in root.iter("item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        summary = re.sub(r"<[^>]+>", "", (item.findtext("description") or "").strip())
        pub_date = (item.findtext("pubDate") or "").strip()
        guid = (item.findtext("guid") or link).strip()
        if title:
            items.append({"title": title, "link": link, "summary": summary[:300],
                          "pub_date": pub_date, "guid": guid})
    # Atom
    ns = "{http://www.w3.org/2005/Atom}"
    for entry in root.iter(f"{ns}entry"):
        title_el = entry.find(f"{ns}title")
        link_el = entry.find(f"{ns}link")
        summary_el = entry.find(f"{ns}summary") or entry.find(f"{ns}content")
        pub_el = entry.find(f"{ns}published") or entry.find(f"{ns}updated")
        id_el = entry.find(f"{ns}id")
        title = ((title_el.text or "") if title_el is not None else "").strip()
        link = ""
        if link_el is not None:
            link = link_el.get("href") or (link_el.text or "")
        summary = ((summary_el.text or "") if summary_el is not None else "").strip()
        summary = re.sub(r"<[^>]+>", "", summary)
        pub_date = ((pub_el.text or "") if pub_el is not None else "").strip()
        guid = ((id_el.text or link) if id_el is not None else link).strip()
        if title:
            items.append({"title": title, "link": link, "summary": summary[:300],
                          "pub_date": pub_date, "guid": guid})
    return items


def matches_challenger(text: str) -> tuple[str | None, str | None]:
    """텍스트 안에 어떤 챌린저 키워드가 매칭되는지 반환."""
    if not text:
        return None, None
    text_low = text.lower()
    for c in CHALLENGERS:
        for kw in c["rss_keywords"]:
            if kw.lower() in text_low:
                return c["id"], kw
    return None, None


# ====================================================================
# 상장 챌린저 yfinance 데이터
# ====================================================================
def fetch_market_data(ticker: str) -> dict:
    """상장 챌린저의 현재가/시총."""
    try:
        t = yf.Ticker(ticker)
        info = t.info or {}
        return {
            "current_price": info.get("currentPrice") or info.get("regularMarketPrice"),
            "market_cap": info.get("marketCap"),
            "currency": info.get("currency"),
            "change_pct": info.get("regularMarketChangePercent"),
        }
    except Exception:
        return {}


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 72)
    print("  반도체 챌린저 모니터 (신흥 위협 분석)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  대상: 챌린저 {len(CHALLENGERS)}개 · 기존사 {len(INCUMBENTS)}개")
    print("=" * 72)

    state = load_state(STATE_NAME, default={"alerted_news_guids": [], "translations": {}})
    alerted_guids = set(state.get("alerted_news_guids", []))
    translation_cache = state.get("translations", {})

    # 1) 상장 챌린저 시장 데이터
    print("\n[상장 챌린저 시장 데이터]")
    market_data = {}
    for c in CHALLENGERS:
        if c.get("ticker"):
            md = fetch_market_data(c["ticker"])
            if md:
                market_data[c["id"]] = md
                print(f"  {c['name']:30s} ({c['ticker']}) | 가격 {md.get('current_price')} {md.get('currency', '')}")

    # 2) RSS 뉴스 매칭 + 번역
    print("\n[RSS 뉴스 매칭 + 한국어 번역]")
    news_by_challenger: dict = {c["id"]: [] for c in CHALLENGERS}
    new_alerts = []
    total_matched = 0

    for feed in RSS_FEEDS:
        print(f"  RSS: {feed['name']}...")
        try:
            raw = http_get(feed["url"])
        except Exception as e:
            print(f"    [실패] {e}")
            continue
        items = parse_rss_items(raw)
        for it in items:
            haystack = f"{it['title']} {it['summary']}"
            cid, kw = matches_challenger(haystack)
            if not cid:
                continue
            it["source"] = feed["name"]
            it["matched_keyword"] = kw
            it["challenger_id"] = cid

            # 번역 (한국어 매체는 번역 불필요)
            guid = it["guid"]
            if feed["lang"] == "ko":
                it["title_ko"] = it["title"]
                it["summary_ko"] = it["summary"]
            elif guid in translation_cache:
                it["title_ko"] = translation_cache[guid].get("title_ko")
                it["summary_ko"] = translation_cache[guid].get("summary_ko")
            else:
                title_ko = translate_to_korean(it["title"])
                summary_ko = translate_to_korean(it["summary"]) if it.get("summary") else None
                it["title_ko"] = title_ko
                it["summary_ko"] = summary_ko
                translation_cache[guid] = {"title_ko": title_ko, "summary_ko": summary_ko}

            news_by_challenger[cid].append(it)
            total_matched += 1

            if guid not in alerted_guids:
                new_alerts.append(it)
                alerted_guids.add(guid)
        print(f"    총 {len(items)}건 중 챌린저 매칭 누적 {total_matched}건")

    # 3) 챌린저 카드 조립
    print("\n[챌린저별 영향 분석 카드 조립]")
    challengers_out = []
    for c in CHALLENGERS:
        cid = c["id"]
        # impact_thesis에 incumbent 이름 매핑
        impacts = []
        for inc_id, thesis in (c.get("impact_thesis") or {}).items():
            inc_info = INCUMBENTS.get(inc_id, {"name": inc_id})
            impacts.append({
                "incumbent_id": inc_id,
                "incumbent_name": inc_info.get("name"),
                "incumbent_role": inc_info.get("role"),
                "sentiment": thesis.get("sentiment"),
                "magnitude": thesis.get("magnitude"),
                "reasoning": thesis.get("reasoning"),
            })

        # 뉴스 최신순 (pub_date 파싱 어려워 그냥 추가 순)
        news = news_by_challenger.get(cid, [])[:5]

        challengers_out.append({
            **{k: v for k, v in c.items() if k not in ("impact_thesis", "rss_keywords")},
            "market": market_data.get(cid),
            "impacts": impacts,
            "recent_news": news,
            "news_count": len(news_by_challenger.get(cid, [])),
        })
        if news:
            print(f"  {c['name']:30s} 뉴스 {len(news_by_challenger[cid])}건")

    # 4) 텔레그램 알림 (신규 뉴스만)
    if new_alerts:
        lines = ["🔬 반도체 챌린저 신규 뉴스", "=" * 30, ""]
        by_challenger: dict = {}
        for n in new_alerts:
            by_challenger.setdefault(n["challenger_id"], []).append(n)
        for cid, items in list(by_challenger.items())[:5]:
            c = next((x for x in CHALLENGERS if x["id"] == cid), None)
            if not c:
                continue
            lines.append(f"🆕 {c['name']} ({c['category']})")
            for n in items[:3]:
                title = n.get("title_ko") or n.get("title", "")
                lines.append(f"  📰 [{n['source']}] {title[:80]}")
                if n.get("link"):
                    lines.append(f"     {n['link']}")
            lines.append("")
        lines.append("🚨 신흥 위협 모니터. 자동 매매 금지.")
        lines.append("대시보드: https://15678910.github.io/ai-finance/")
        try:
            send_message("\n".join(lines))
            print(f"\n  ✅ 텔레그램 발송: 신규 {len(new_alerts)}건")
        except Exception as e:
            print(f"\n  ❌ 텔레그램 발송 실패: {e}")

    # 5) 상태 + 결과 저장
    state["alerted_news_guids"] = list(alerted_guids)[-500:]
    if len(translation_cache) > 500:
        recent_keys = list(translation_cache.keys())[-500:]
        translation_cache = {k: translation_cache[k] for k in recent_keys}
    state["translations"] = translation_cache
    save_state(STATE_NAME, state)

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "methodology": (
            "Phase 1 (Heuristic): 사람이 사전 정의한 영향 thesis + RSS 키워드 매칭 + 자동 한국어 번역. "
            "분기마다 thesis 수동 업데이트 권장. Phase 2 (LLM-enhanced) 후속 계획."
        ),
        "challengers_count": len(challengers_out),
        "incumbents": [{"id": k, **v} for k, v in INCUMBENTS.items()],
        "challengers": challengers_out,
        "total_news_matched": total_matched,
        "new_alerts_this_run": len(new_alerts),
        "data_sources": [f["name"] for f in RSS_FEEDS] + ["yfinance"] + ["Google Translate"],
        "warning": "🚨 휴리스틱 분석. 실거래 단독 사용 금지. thesis는 분기마다 검토 필요.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    print("\n" + "=" * 72)
    print(f"  완료: 챌린저 {len(challengers_out)}개 · 뉴스 매칭 {total_matched}건 (신규 {len(new_alerts)})")
    print("=" * 72)


if __name__ == "__main__":
    main()
