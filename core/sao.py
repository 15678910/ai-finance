"""
Standard Asset Object (SAO) — 통합 자산 데이터 모델
=====================================================

목적:
  모든 분석 모듈이 공통으로 출력할 수 있는 표준 시그널/엔티티 형식.
  Sovereign Watch의 TAK Protocol에서 영감을 받았으며, 금융 도메인에
  맞게 자체 정의.

  → docs/sao_signals.json 으로 직렬화
  → 대시보드의 Event Feed가 이 형식을 읽음

설계 원칙:
  · 도메인 비종속: Stock/Index/Bill/MacroEvent 등 다양한 자산 표현
  · 출처 명시: 모든 데이터에 source 필드 (데이터 투명성)
  · 우선순위 자동: P1(즉시)/P2(주의)/P3(관찰)/P4(정보)
  · KG 호환: relations 필드로 지식 그래프와 연동

🚨 시뮬레이션. 자동 매매 금지.
"""

from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import Optional
from datetime import datetime, timezone, timedelta


# ====================================================================
# 분류 상수
# ====================================================================
# 6개 Financial Pulse (Sovereign Watch의 Pulse 컨셉)
PULSE_TYPES = {
    "MACRO":      {"emoji": "🏦", "color": "#ef4444", "label": "거시·신용"},
    "KR_EQUITY":  {"emoji": "🇰🇷", "color": "#22c55e", "label": "한국 종목"},
    "CRYPTO":     {"emoji": "🪙", "color": "#f97316", "label": "암호화폐"},
    "LEGISLATIVE": {"emoji": "🏛️", "color": "#fb923c", "label": "입법"},
    "SEMI":       {"emoji": "🔬", "color": "#06b6d4", "label": "반도체"},
    "NEWS":       {"emoji": "📰", "color": "#9ca3af", "label": "뉴스"},
}

# 엔티티 타입
ENTITY_TYPES = ["Stock", "Index", "Currency", "Bond", "Bill", "NewsItem",
                "MacroEvent", "Sector", "Company"]

# 우선순위 (P1: 즉시 행동, P2: 주의, P3: 관찰, P4: 정보)
PRIORITIES = {
    "P1": {"emoji": "🚨", "color": "#dc2626", "label": "Critical (즉시 행동)",      "action": "Act"},
    "P2": {"emoji": "⚠️", "color": "#f97316", "label": "Alert (주의)",            "action": "Watch"},
    "P3": {"emoji": "🟡", "color": "#fbbf24", "label": "Watch (관찰)",            "action": "Note"},
    "P4": {"emoji": "🟢", "color": "#94a3b8", "label": "Info (정보)",             "action": "Read"},
}


# ====================================================================
# 데이터 클래스
# ====================================================================
@dataclass
class Signal:
    """단일 활성 시그널 (예: '외인 5일 연속 매도')."""
    type: str                # "consensus_buy", "yield_curve_inverted" 등
    strength: int            # 1~5
    emoji: str
    label: str
    description: str = ""
    detected_at: Optional[str] = None    # ISO 8601


@dataclass
class Relation:
    """KG 호환 관계 (다른 SAO 객체 참조)."""
    type: str                # "IN_SECTOR", "CORRELATES_WITH", "AFFECTED_BY" 등
    target_id: str           # 다른 SAO id
    weight: float = 1.0
    data: dict = field(default_factory=dict)


@dataclass
class AssetObject:
    """Standard Asset Object — 모든 모듈의 출력 표준."""
    # ─── 식별 ─────────────────────────────────
    id: str                                # "stock:005930", "fred:DGS10", "bill:HR3633"
    type: str                              # ENTITY_TYPES 중 하나
    name: str                              # 화면 표시명 (예: "삼성전자")

    # ─── 분류 ─────────────────────────────────
    pulse: str                             # PULSE_TYPES 중 하나
    market: Optional[str] = None           # "KOSPI" / "KOSDAQ" / "NYSE" / "NASDAQ" / "KRX_DEBT"
    sector: Optional[str] = None           # "IT/반도체" 등

    # ─── 메타데이터 ───────────────────────────
    source: list = field(default_factory=list)    # ["FRED", "yfinance"]
    last_updated: Optional[str] = None     # ISO 8601 KST

    # ─── 핵심 값 ──────────────────────────────
    value: Optional[float] = None
    unit: str = ""

    # ─── 우선순위 + 시그널 ────────────────────
    priority: str = "P4"                   # "P1"/"P2"/"P3"/"P4"
    signals: list = field(default_factory=list)   # list[Signal]

    # ─── 메트릭 (자유 형식) ───────────────────
    metrics: dict = field(default_factory=dict)

    # ─── 관계 (KG 호환) ───────────────────────
    relations: list = field(default_factory=list)   # list[Relation]

    # ─── 부가 정보 ────────────────────────────
    description: str = ""
    link: Optional[str] = None             # 원본 페이지 (FRED URL 등)
    icon: Optional[str] = None             # 이모지 또는 아이콘 식별자

    def to_dict(self) -> dict:
        """JSON 직렬화."""
        return asdict(self)


# ====================================================================
# Helper Functions
# ====================================================================
def now_kst_iso() -> str:
    """현재 시각 (KST, ISO 8601)."""
    kst = timezone(timedelta(hours=9))
    return datetime.now(kst).isoformat(timespec="seconds")


def stock_id(ticker: str) -> str:
    """종목 ID 정규화. '005930.KS' → 'stock:005930'."""
    if not ticker:
        return ""
    code = ticker.replace(".KS", "").replace(".KQ", "")
    return f"stock:{code}"


def fred_id(series_id: str) -> str:
    return f"fred:{series_id}"


def bill_id(short_name: str) -> str:
    import re
    slug = re.sub(r"\s+", "", short_name or "")
    return f"bill:{slug}"


def news_id(guid: str) -> str:
    return f"news:{abs(hash(guid)) % 10**12}"


def sector_id(name: str) -> str:
    import re
    slug = re.sub(r"[\s/]+", "", name or "")
    return f"sector:{slug}"


def currency_id(code: str) -> str:
    return f"currency:{code}"


def index_id(ticker: str) -> str:
    return f"index:{ticker}"


# ====================================================================
# 우선순위 자동 분류 규칙
# ====================================================================
def classify_priority(signals: list) -> str:
    """시그널 리스트에서 가장 강한 우선순위 결정.

    strength 5 → P1
    strength 4 → P2
    strength 3 → P3
    strength 1~2 또는 없음 → P4
    """
    if not signals:
        return "P4"
    max_strength = max((s.get("strength", 0) if isinstance(s, dict) else getattr(s, "strength", 0))
                       for s in signals)
    if max_strength >= 5:
        return "P1"
    elif max_strength >= 4:
        return "P2"
    elif max_strength >= 3:
        return "P3"
    return "P4"


def make_signal(type_: str, strength: int, emoji: str, label: str,
                description: str = "") -> dict:
    """Signal 객체를 dict로 빠르게 생성."""
    return {
        "type": type_,
        "strength": strength,
        "emoji": emoji,
        "label": label,
        "description": description,
        "detected_at": now_kst_iso(),
    }


# ====================================================================
# 직렬화 컨테이너 (대시보드 출력용)
# ====================================================================
def build_signals_payload(objects: list) -> dict:
    """모든 SAO 객체를 우선순위로 정렬 + Event Feed 컨테이너 형식 생성."""
    # AssetObject → dict 변환
    items = []
    for obj in objects:
        if isinstance(obj, AssetObject):
            items.append(obj.to_dict())
        elif isinstance(obj, dict):
            items.append(obj)

    # 우선순위 정렬 (P1 → P4)
    priority_order = {"P1": 0, "P2": 1, "P3": 2, "P4": 3}
    items.sort(key=lambda x: (priority_order.get(x.get("priority", "P4"), 4),
                              -max((s.get("strength", 0) for s in x.get("signals", [])), default=0)))

    # 카운트
    counts = {p: 0 for p in PRIORITIES}
    for item in items:
        counts[item.get("priority", "P4")] = counts.get(item.get("priority", "P4"), 0) + 1

    # Pulse별 카운트
    pulse_counts = {p: 0 for p in PULSE_TYPES}
    for item in items:
        pulse_counts[item.get("pulse", "MACRO")] = pulse_counts.get(item.get("pulse", "MACRO"), 0) + 1

    return {
        "generated_at": now_kst_iso(),
        "total": len(items),
        "priority_counts": counts,
        "pulse_counts": pulse_counts,
        "ready_to_act": counts.get("P1", 0) + counts.get("P2", 0),  # Sovereign Watch "Ready to Task"
        "pulse_metadata": PULSE_TYPES,
        "priority_metadata": PRIORITIES,
        "signals": items,
        "warning": "🚨 통합 시그널 큐. 정보 제공 목적. 자동 매매 금지.",
    }
