"""
통합 지식 그래프 (Knowledge Graph) — 팔란티어 스타일 객체 중심 분석
=========================================================================

목적:
  ai-finance의 모든 분석 모듈(IRR·반도체민감도·DCF·CLARITY·저평가스크리너 등)에서
  생성된 데이터를 단일 그래프로 통합하여 종목·섹터·법안·지수·통화 간의
  **의미적 연결**을 탐색 가능하게 함.

핵심 개념:
  - Entity: 종목/섹터/지수/법안/통화/거시이벤트 (각각 고유 URI)
  - Relation: 정형 관계 (IN_SECTOR, CORRELATES_WITH, AFFECTED_BY 등)
  - 그래프 = 노드 + 엣지의 단일 직렬화 (vis.js 호환)

설계 원칙:
  - 외부 의존성 없음 (Python 표준 + dataclass만)
  - JSON 직렬화 가능 (대시보드로 그대로 전달)
  - 엔티티 정규화: 동일 종목이 다른 모듈에서 다른 ID로 등장해도 통합

ID 컨벤션:
  stock:005930 / index:NQ=F / sector:IT반도체 / bill:HR3633 / currency:KRW
"""

from __future__ import annotations
from dataclasses import dataclass, field
import re


# ====================================================================
# 엔티티 타입 정의
# ====================================================================
ENTITY_TYPES = {
    "Stock":      {"emoji": "📈", "color": "#22d3ee"},   # cyan
    "Sector":     {"emoji": "🏢", "color": "#a78bfa"},   # purple
    "Index":      {"emoji": "📊", "color": "#fbbf24"},   # amber
    "Bill":       {"emoji": "🏛️", "color": "#fb923c"},   # orange
    "Currency":   {"emoji": "💱", "color": "#4ade80"},   # green
    "MacroEvent": {"emoji": "🌐", "color": "#f87171"},   # red
    "NewsItem":   {"emoji": "📰", "color": "#9ca3af"},   # gray
}

# 관계 타입 정의 (방향성 있음)
RELATION_TYPES = {
    "IN_SECTOR":       {"emoji": "🔸", "weight": 1.0, "description": "종목이 속한 섹터"},
    "CORRELATES_WITH": {"emoji": "↔️", "weight": 1.0, "description": "지수와 통계적 상관 (β + R²)"},
    "SENSITIVE_TO":    {"emoji": "💧", "weight": 0.8, "description": "통화/거시 변수에 민감"},
    "AFFECTED_BY":     {"emoji": "⚡", "weight": 0.9, "description": "법안/이벤트의 영향"},
    "MENTIONS":        {"emoji": "🔗", "weight": 0.5, "description": "뉴스가 엔티티를 언급"},
    "COMPETES_WITH":   {"emoji": "⚔️", "weight": 0.7, "description": "동일 시장 경쟁사"},
    "SUPPLIES":        {"emoji": "🚚", "weight": 0.6, "description": "공급망 관계"},
}


# ====================================================================
# 노드/엣지 데이터 클래스
# ====================================================================
@dataclass
class Node:
    """그래프 노드 (엔티티)."""
    id: str               # 고유 URI (예: "stock:005930")
    type: str             # ENTITY_TYPES key
    label: str            # 화면 표시명 (예: "삼성전자")
    data: dict = field(default_factory=dict)   # 모듈별 데이터 (metrics, IRR 등 통합)

    def merge(self, other_data: dict):
        """기존 데이터에 새 모듈 데이터를 병합 (덮어쓰지 않고 보완)."""
        for k, v in other_data.items():
            if k not in self.data or self.data[k] is None:
                self.data[k] = v
            elif isinstance(self.data[k], dict) and isinstance(v, dict):
                # 중첩 dict는 재귀 병합 (예: metrics 안의 새 키 추가)
                self.data[k] = {**self.data[k], **v}


@dataclass
class Edge:
    """그래프 엣지 (관계)."""
    source: str           # from node ID
    target: str           # to node ID
    type: str             # RELATION_TYPES key
    weight: float = 1.0   # 시각화 두께
    data: dict = field(default_factory=dict)  # 관계의 부가 정보 (β, R² 등)


# ====================================================================
# 메인 KG 클래스
# ====================================================================
class KnowledgeGraph:
    def __init__(self):
        self.nodes: dict[str, Node] = {}
        self.edges: list[Edge] = []
        self._edge_set: set = set()  # (source, target, type) 중복 방지

    # ─── 노드 관리 ──────────────────────────────────────
    def add_node(self, node_id: str, node_type: str, label: str, data: dict | None = None) -> Node:
        """노드 추가 또는 병합."""
        if node_id in self.nodes:
            if data:
                self.nodes[node_id].merge(data)
            return self.nodes[node_id]
        node = Node(id=node_id, type=node_type, label=label, data=data or {})
        self.nodes[node_id] = node
        return node

    def get_node(self, node_id: str) -> Node | None:
        return self.nodes.get(node_id)

    # ─── 엣지 관리 ──────────────────────────────────────
    def add_edge(self, source: str, target: str, rel_type: str,
                 weight: float = 1.0, data: dict | None = None) -> bool:
        """엣지 추가. 양 끝 노드가 모두 존재해야 함. 중복 방지."""
        if source not in self.nodes or target not in self.nodes:
            return False
        key = (source, target, rel_type)
        if key in self._edge_set:
            # 동일 엣지 존재 시 weight·data 업데이트
            for e in self.edges:
                if e.source == source and e.target == target and e.type == rel_type:
                    e.weight = max(e.weight, weight)
                    if data:
                        e.data = {**e.data, **data}
                    return False
        self.edges.append(Edge(source=source, target=target, type=rel_type, weight=weight, data=data or {}))
        self._edge_set.add(key)
        return True

    # ─── ID 정규화 헬퍼 ─────────────────────────────────
    @staticmethod
    def stock_id(ticker_or_code: str) -> str:
        """종목 ID 정규화. '005930.KS' → 'stock:005930', '삼성전자' 같은 이름은 별도 처리."""
        if not ticker_or_code:
            return ""
        # ticker에서 .KS/.KQ 제거
        code = ticker_or_code.replace(".KS", "").replace(".KQ", "")
        return f"stock:{code}"

    @staticmethod
    def sector_id(sector_name: str) -> str:
        if not sector_name:
            return ""
        # 공백·슬래시 제거하여 URI safe
        slug = re.sub(r"[\s/]+", "", sector_name)
        return f"sector:{slug}"

    @staticmethod
    def index_id(ticker: str) -> str:
        return f"index:{ticker}"

    @staticmethod
    def bill_id(bill_short: str) -> str:
        slug = re.sub(r"\s+", "", bill_short)
        return f"bill:{slug}"

    @staticmethod
    def currency_id(code: str) -> str:
        return f"currency:{code}"

    @staticmethod
    def macro_id(event: str) -> str:
        slug = re.sub(r"\s+", "_", event)
        return f"macro:{slug}"

    # ─── 직렬화 ────────────────────────────────────────
    def to_vis_format(self) -> dict:
        """vis.js network 형식으로 변환 (대시보드용)."""
        nodes_out = []
        for n in self.nodes.values():
            etype = ENTITY_TYPES.get(n.type, {})
            nodes_out.append({
                "id": n.id,
                "label": n.label,
                "type": n.type,
                "color": etype.get("color", "#888"),
                "emoji": etype.get("emoji", "•"),
                "data": n.data,
            })

        edges_out = []
        for e in self.edges:
            edges_out.append({
                "from": e.source,
                "to": e.target,
                "type": e.type,
                "weight": e.weight,
                "data": e.data,
            })

        return {
            "nodes": nodes_out,
            "edges": edges_out,
            "stats": {
                "node_count": len(nodes_out),
                "edge_count": len(edges_out),
                "node_types": {t: sum(1 for n in self.nodes.values() if n.type == t)
                               for t in ENTITY_TYPES},
                "edge_types": {t: sum(1 for e in self.edges if e.type == t)
                               for t in RELATION_TYPES},
            },
        }

    def stats_summary(self) -> str:
        v = self.to_vis_format()
        s = v["stats"]
        lines = [f"노드 {s['node_count']} · 엣지 {s['edge_count']}", ""]
        lines.append("[엔티티 타입별]")
        for t, c in s["node_types"].items():
            if c > 0:
                lines.append(f"  {ENTITY_TYPES[t]['emoji']} {t}: {c}")
        lines.append("[관계 타입별]")
        for t, c in s["edge_types"].items():
            if c > 0:
                lines.append(f"  {RELATION_TYPES[t]['emoji']} {t}: {c}")
        return "\n".join(lines)
