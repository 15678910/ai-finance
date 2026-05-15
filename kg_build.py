"""
지식 그래프 빌더
================

모든 분석 모듈의 JSON을 읽어 통합 지식 그래프를 생성.

입력 JSON:
  - docs/value_screener.json        (종목 메트릭 + IRR 컬럼)
  - docs/irr_analysis.json          (다중 IRR + 안전마진)
  - docs/semi_sensitivity.json      (단변량 β NQ/SOX)
  - docs/semi_sensitivity_v2.json   (다변량 β + KOSPI/USDKRW)
  - docs/dcf_valuations.json        (DCF 적정가)
  - docs/clarity_act.json           (법안 상태 + 뉴스)
  - docs/bitcoin_standard.json      (BTC 매크로)
  - docs/breaking_news_state.json   (긴급 뉴스)

출력:
  - docs/knowledge_graph.json   (vis.js 호환 노드+엣지)

이 모듈은 cron 워크플로의 마지막 단계로 실행되어
다른 모듈들의 최신 결과를 통합 그래프로 합성합니다.

🚨 시뮬레이션. 자동 매매 금지.
"""

import os
import json
from datetime import datetime, timezone, timedelta

from core.knowledge_graph import KnowledgeGraph

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS_DIR = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS_DIR, "knowledge_graph.json")
KST = timezone(timedelta(hours=9))


def _load_json(name: str) -> dict | None:
    """안전한 JSON 로드 (없으면 None)."""
    path = os.path.join(DOCS_DIR, name)
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"  [로드 실패] {name}: {e}")
        return None


# ====================================================================
# 빌드 단계별 함수
# ====================================================================
def build_from_value_screener(kg: KnowledgeGraph):
    """저평가 스크리너 → 종목 + 섹터 + IRR + 점수."""
    data = _load_json("value_screener.json")
    if not data:
        print("  [건너뜀] value_screener.json 없음")
        return

    stocks = (data.get("top_picks", []) or []) + (data.get("watchlist", []) or [])
    for s in stocks:
        sid = KnowledgeGraph.stock_id(s.get("ticker", ""))
        if not sid:
            continue

        kg.add_node(sid, "Stock", s.get("name", ""), {
            "ticker": s.get("ticker"),
            "name": s.get("name"),
            "market": s.get("market"),
            "sector": s.get("sector"),
            "current_price": s.get("current_price"),
            "market_cap_billion": s.get("market_cap_billion"),
            "metrics": s.get("metrics", {}),
            "value_score": (s.get("scores", {}) or {}).get("total_score"),
            "irr_basic": s.get("irr", {}),  # Phase C IRR
            "filter_passed": s.get("filter_passed"),
        })

        # 섹터 노드 + IN_SECTOR 엣지
        sector = s.get("sector")
        if sector:
            sec_id = KnowledgeGraph.sector_id(sector)
            kg.add_node(sec_id, "Sector", sector, {"name": sector})
            kg.add_edge(sid, sec_id, "IN_SECTOR")
    print(f"  · value_screener 종목 {len(stocks)}개 통합")


def build_from_irr_analysis(kg: KnowledgeGraph):
    """IRR 분석기 → 종목에 다중 IRR + 안전마진 + 등급 추가."""
    data = _load_json("irr_analysis.json")
    if not data:
        print("  [건너뜀] irr_analysis.json 없음")
        return

    for r in data.get("results", []):
        sid = KnowledgeGraph.stock_id(r.get("ticker", ""))
        if not sid:
            continue
        # 종목이 아직 없으면 신규 추가, 있으면 데이터 병합
        kg.add_node(sid, "Stock", r.get("name", ""), {
            "ticker": r.get("ticker"),
            "name": r.get("name"),
            "sector": r.get("sector"),
            "market": r.get("market"),
            "current_price": r.get("current_price"),
            "irr_full": r.get("irr"),          # 3가지 IRR 방법
            "required_return": r.get("required_return"),
            "margin_of_safety_pct": r.get("margin_of_safety_pct"),
            "rating": r.get("rating"),
            "rating_color": r.get("rating_color"),
            "is_strong_buy": r.get("is_strong_buy"),
            "snapshot": r.get("snapshot"),
        })
        # 섹터 엣지
        sector = r.get("sector")
        if sector:
            sec_id = KnowledgeGraph.sector_id(sector)
            kg.add_node(sec_id, "Sector", sector, {"name": sector})
            kg.add_edge(sid, sec_id, "IN_SECTOR")
    print(f"  · irr_analysis 종목 {len(data.get('results', []))}개 통합 + 섹터 평균 IRR")

    # 섹터별 평균 IRR을 sector 노드에 추가
    for s in data.get("sector_summary", []):
        sec_id = KnowledgeGraph.sector_id(s.get("sector", ""))
        if sec_id in kg.nodes:
            kg.nodes[sec_id].data["mean_irr_pct"] = s.get("mean_irr_pct")
            kg.nodes[sec_id].data["median_irr_pct"] = s.get("median_irr_pct")
            kg.nodes[sec_id].data["count"] = s.get("count")


def build_from_semi_sensitivity(kg: KnowledgeGraph):
    """반도체 민감도 v1 (단변량) → 종목 ↔ 지수 상관 엣지."""
    data = _load_json("semi_sensitivity.json")
    if not data:
        print("  [건너뜀] semi_sensitivity.json 없음")
        return

    # 미국 벤치마크 노드 추가
    for b in data.get("us_benchmarks", []):
        idx_id = KnowledgeGraph.index_id(b.get("ticker", ""))
        kg.add_node(idx_id, "Index", b.get("name", b.get("ticker", "")), {
            "ticker": b.get("ticker"),
            "key": b.get("key"),
        })

    # 한국 종목 ↔ 미국 지수 상관 엣지
    edge_count = 0
    for stock in data.get("results", []):
        sid = KnowledgeGraph.stock_id(stock.get("ticker", ""))
        if sid not in kg.nodes:
            # 종목 노드 신규 생성 (value_screener에 없을 경우)
            kg.add_node(sid, "Stock", stock.get("name", ""), {
                "ticker": stock.get("ticker"),
                "name": stock.get("name"),
            })
        kg.nodes[sid].data["semi_betas_v1"] = stock.get("benchmarks", {})

        # 각 벤치마크별 엣지
        for bm_key, bm in (stock.get("benchmarks") or {}).items():
            if not bm or bm.get("beta") is None:
                continue
            idx_id = KnowledgeGraph.index_id(bm.get("benchmark_ticker", ""))
            if idx_id not in kg.nodes:
                continue
            beta = bm.get("beta")
            r2 = bm.get("r_squared") or 0
            # 의미 있는 상관만 표시 (R² > 0.1 또는 |β| > 0.5)
            if (r2 and r2 > 0.10) or (beta and abs(beta) > 0.5):
                kg.add_edge(sid, idx_id, "CORRELATES_WITH", weight=abs(beta), data={
                    "beta": beta,
                    "r_squared": r2,
                    "method": "단변량 OLS (v1, T-1 lag)",
                })
                edge_count += 1
    print(f"  · semi_sensitivity v1 → 종목↔지수 상관 엣지 {edge_count}개")


def build_from_semi_sensitivity_v2(kg: KnowledgeGraph):
    """반도체 민감도 v2 (다변량) → 종목에 v2 결과 + USDKRW 민감도 엣지."""
    data = _load_json("semi_sensitivity_v2.json")
    if not data:
        print("  [건너뜀] semi_sensitivity_v2.json 없음")
        return

    # 한국 변수 (USDKRW, KOSPI) 노드 추가
    for v in (data.get("variable_definitions", {}) or {}).get("kr_lag0", []):
        if v.get("key") == "usdkrw":
            kg.add_node(KnowledgeGraph.currency_id("KRW"), "Currency", "USD/KRW", {
                "ticker": v.get("ticker"), "key": v.get("key"),
            })
        elif v.get("key") == "kospi":
            kg.add_node(KnowledgeGraph.index_id("^KS11"), "Index", "KOSPI", {
                "ticker": v.get("ticker"), "key": v.get("key"),
            })
    # US 10Y 금리도 거시 이벤트 노드로
    for v in (data.get("variable_definitions", {}) or {}).get("us_lag1", []):
        if v.get("key") == "ust10y":
            kg.add_node(KnowledgeGraph.macro_id("UST10Y"), "MacroEvent", "US 10Y 금리", {
                "ticker": v.get("ticker"),
            })

    # v2 다변량 결과 종목에 추가 + KOSPI/USDKRW 엣지
    for stock in data.get("results", []):
        sid = KnowledgeGraph.stock_id(stock.get("ticker", ""))
        if sid not in kg.nodes:
            kg.add_node(sid, "Stock", stock.get("name", ""), {
                "ticker": stock.get("ticker"),
                "name": stock.get("name"),
            })
        kg.nodes[sid].data["semi_v2"] = stock.get("multivariate", {})

        betas = (stock.get("multivariate") or {}).get("betas", {})
        # KOSPI 엣지
        if betas.get("kospi") is not None:
            kg.add_edge(sid, KnowledgeGraph.index_id("^KS11"), "CORRELATES_WITH",
                        weight=abs(betas["kospi"]),
                        data={"beta_kospi": betas["kospi"], "method": "다변량 OLS v2"})
        # USDKRW 엣지
        if betas.get("usdkrw") is not None and abs(betas["usdkrw"]) > 0.1:
            kg.add_edge(sid, KnowledgeGraph.currency_id("KRW"), "SENSITIVE_TO",
                        weight=abs(betas["usdkrw"]),
                        data={"beta_usdkrw": betas["usdkrw"], "method": "다변량 OLS v2"})
    print(f"  · semi_sensitivity v2 다변량 결과 종목 {len(data.get('results', []))}개 통합")


def build_from_dcf(kg: KnowledgeGraph):
    """DCF 평가 → 종목에 적정가·업사이드 추가."""
    data = _load_json("dcf_valuations.json")
    if not data:
        print("  [건너뜀] dcf_valuations.json 없음")
        return
    count = 0
    for v in data.get("valuations", []):
        sid = KnowledgeGraph.stock_id(v.get("ticker", ""))
        if not sid:
            continue
        kg.add_node(sid, "Stock", v.get("name", ""), {
            "ticker": v.get("ticker"),
            "name": v.get("name"),
            "dcf": {
                "fair_price": v.get("fair_price_dcf"),
                "upside_pct": v.get("upside_pct"),
                "wacc_pct": v.get("wacc_pct"),
                "growth_rate_pct": v.get("growth_rate_pct"),
                "signal": v.get("signal"),
            }
        })
        count += 1
    print(f"  · DCF 평가 종목 {count}개 통합")


def build_from_clarity_act(kg: KnowledgeGraph):
    """법안 + 뉴스 → Bill 노드 + NewsItem 노드 + AFFECTED_BY 엣지."""
    data = _load_json("clarity_act.json")
    if not data:
        print("  [건너뜀] clarity_act.json 없음")
        return

    # 법안 노드
    bill_count = 0
    for b in data.get("bills", []):
        bid = KnowledgeGraph.bill_id(b.get("short_name", ""))
        if not bid:
            continue
        kg.add_node(bid, "Bill", b.get("short_name", ""), {
            "short_name": b.get("short_name"),
            "full_name": b.get("full_name"),
            "congress": b.get("congress"),
            "bill_type": b.get("bill_type"),
            "number": b.get("number"),
            "status_label_kr": b.get("status_label_kr"),
            "status_date": b.get("status_date"),
            "is_alive": b.get("is_alive"),
            "category": b.get("category"),
            "link": b.get("link"),
            "sponsor": b.get("sponsor"),
        })
        bill_count += 1

    # 암호화폐 입법 → 암호화폐 관련 한국 종목 영향 가정
    # (간단 휴리스틱: 종목 이름·섹터에 '암호화폐'·'블록체인' 포함 — 현재 universe엔 없음)
    # 대신 입법 통과 시 시장 전반에 영향을 줄 수 있음을 KOSPI 엣지로 표현
    kospi_id = KnowledgeGraph.index_id("^KS11")
    if kospi_id in kg.nodes:
        for b in data.get("bills", []):
            bid = KnowledgeGraph.bill_id(b.get("short_name", ""))
            if bid in kg.nodes:
                kg.add_edge(bid, kospi_id, "AFFECTED_BY", weight=0.3,
                            data={"reason": "암호화폐 입법은 시장 위험선호에 영향"})

    # 뉴스 노드 + 법안 ↔ 뉴스 MENTIONS 엣지
    news_count = 0
    for n in (data.get("recent_news") or [])[:10]:  # 너무 많아지지 않게 상위 10건만
        guid = n.get("guid") or n.get("link") or n.get("title", "")[:50]
        if not guid:
            continue
        nid = f"news:{abs(hash(guid)) % 10**12}"
        kg.add_node(nid, "NewsItem", (n.get("title_ko") or n.get("title") or "")[:50], {
            "title": n.get("title"),
            "title_ko": n.get("title_ko"),
            "summary_ko": n.get("summary_ko"),
            "source": n.get("source"),
            "link": n.get("link"),
            "matched_keyword": n.get("matched_keyword"),
        })
        # 매칭 키워드와 가장 가까운 법안과 연결 (heuristic)
        kw = (n.get("matched_keyword") or "").lower()
        for b in data.get("bills", []):
            bid = KnowledgeGraph.bill_id(b.get("short_name", ""))
            bn = (b.get("short_name") or "").lower()
            if bid in kg.nodes and (bn in kw or kw in bn):
                kg.add_edge(nid, bid, "MENTIONS")
        news_count += 1

    print(f"  · CLARITY Act 법안 {bill_count}건 + 뉴스 {news_count}건 통합")


def build_from_bitcoin_standard(kg: KnowledgeGraph):
    """비트코인 본위제 모니터 → BTC 매크로 이벤트 노드 (참고용)."""
    data = _load_json("bitcoin_standard.json")
    if not data:
        return
    bts = data.get("bts_score") or {}
    if not bts:
        return
    btc_id = KnowledgeGraph.macro_id("BitcoinStandard")
    kg.add_node(btc_id, "MacroEvent", "비트코인 본위제", {
        "score": bts.get("score"),
        "level": bts.get("level"),
        "severity": bts.get("severity"),
        "btc_current": (data.get("btc") or {}).get("current"),
    })
    print(f"  · Bitcoin Standard 매크로 이벤트 통합")


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 70)
    print("  지식 그래프 빌더 (Knowledge Graph)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    kg = KnowledgeGraph()

    print("\n[모듈별 데이터 통합]")
    build_from_value_screener(kg)
    build_from_irr_analysis(kg)
    build_from_semi_sensitivity(kg)
    build_from_semi_sensitivity_v2(kg)
    build_from_dcf(kg)
    build_from_clarity_act(kg)
    build_from_bitcoin_standard(kg)

    # 통계
    print("\n[그래프 통계]")
    print(kg.stats_summary())

    # 저장
    vis_data = kg.to_vis_format()
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "methodology": (
            "ai-finance 모듈 JSON 통합 → 노드(엔티티) + 엣지(관계) 그래프. "
            "vis.js network 형식 호환. 클릭 가능한 객체 중심 탐색."
        ),
        "entity_types": list({n.type for n in kg.nodes.values()}),
        "relation_types": list({e.type for e in kg.edges}),
        **vis_data,
        "warning": "🚨 정보 통합 시각화. 투자 결정 단독 사용 금지.",
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n  결과 저장: {OUTPUT_FILE}")
    print(f"\n[완료] 노드 {len(kg.nodes)} · 엣지 {len(kg.edges)}")


if __name__ == "__main__":
    main()
