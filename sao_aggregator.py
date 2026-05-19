"""
SAO Aggregator — Multi-Asset Fusion Center
=================================================

목적:
  모든 분석 모듈의 JSON 출력을 단일 SAO(Standard Asset Object) 형식으로
  통합. 우선순위(P1-P4) 자동 분류 + Event Feed용 직렬화.

입력 (선택적, 있는 것만 읽음):
  - docs/credit_spread.json        → MACRO Pulse
  - docs/investor_flow.json        → KR_EQUITY (외인·기관 시그널)
  - docs/value_screener.json       → KR_EQUITY (가치 점수)
  - docs/irr_analysis.json         → KR_EQUITY (IRR + MoS)
  - docs/semi_sensitivity.json     → SEMI (단변량 β)
  - docs/semi_sensitivity_v2.json  → SEMI (다변량 β)
  - docs/semi_challengers.json     → SEMI (챌린저 위협)
  - docs/clarity_act.json          → LEGISLATIVE
  - docs/bitcoin_standard.json     → CRYPTO
  - docs/breaking_news_state.json  → NEWS

출력:
  - docs/sao_signals.json  (Event Feed 컨테이너)

🚨 시뮬레이션. 자동 매매 금지.
"""

import os
import json
from datetime import datetime, timezone, timedelta

from core.sao import (
    AssetObject, make_signal, classify_priority, build_signals_payload,
    stock_id, fred_id, bill_id, news_id, currency_id, index_id, now_kst_iso,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS_DIR = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS_DIR, "sao_signals.json")
KST = timezone(timedelta(hours=9))


def _load_json(name: str) -> dict | None:
    path = os.path.join(DOCS_DIR, name)
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"  [실패] {name}: {e}")
        return None


# ====================================================================
# 모듈별 → SAO 변환 함수
# ====================================================================
def from_credit_spread() -> list:
    """credit_spread.json → MACRO Pulse SAOs."""
    data = _load_json("credit_spread.json")
    if not data:
        return []
    objs = []
    for r in data.get("results", []):
        # 기존 시그널을 SAO Signal 포맷으로 변환
        signals = []
        for s in r.get("signals", []):
            signals.append(make_signal(
                type_=s.get("type", "unknown"),
                strength=s.get("strength", 3),
                emoji=s.get("emoji", "📋"),
                label=s.get("label", ""),
                description=s.get("description", ""),
            ))

        sao = AssetObject(
            id=fred_id(r["id"]),
            type="MacroEvent" if r["category"] == "yield" else "Bond",
            name=r["name"],
            pulse="MACRO",
            source=["FRED"],
            last_updated=data.get("generated_at"),
            value=r.get("latest_value"),
            unit=r.get("unit", ""),
            signals=[s for s in signals],
            metrics={
                "change_1d": r.get("change_1d"),
                "change_5d": r.get("change_5d"),
                "change_30d": r.get("change_30d"),
                "percentile_1y": r.get("percentile_1y"),
                "avg_1y": r.get("avg_1y"),
            },
            description=r.get("description", ""),
            link=r.get("fred_url"),
            icon="🔴" if r["category"] == "credit_spread" else "💰" if r["category"] == "yield" else "📊",
        )
        sao.priority = classify_priority(signals)
        objs.append(sao)
    print(f"  · credit_spread: {len(objs)}개 SAO")
    return objs


def from_investor_flow() -> list:
    """investor_flow.json → KR_EQUITY Pulse SAOs."""
    data = _load_json("investor_flow.json")
    if not data:
        return []
    objs = []
    for r in data.get("results", []):
        # 강력 시그널만 SAO로 변환 (max_strength >= 3)
        max_strength = r.get("max_strength", 0)
        if max_strength < 3:
            continue
        signals = []
        for s in r.get("signals", []):
            signals.append(make_signal(
                type_=s.get("type", ""),
                strength=s.get("strength", 3),
                emoji=s.get("emoji", "📋"),
                label=s.get("label", ""),
                description=s.get("description", ""),
            ))
        sao = AssetObject(
            id=stock_id(r["ticker"]),
            type="Stock",
            name=r["name"],
            pulse="KR_EQUITY",
            market="KOSPI" if not r["ticker"].endswith(("KQ",)) else "KOSDAQ",
            sector=r.get("sector"),
            source=["Naver 금융"],
            last_updated=data.get("generated_at"),
            value=r.get("latest_close"),
            unit="원",
            signals=signals,
            metrics={
                "foreign_5d_net_shares": r.get("foreign_5d_net_shares"),
                "institutional_5d_net_shares": r.get("institutional_5d_net_shares"),
                "foreign_holding_pct": r.get("foreign_holding_pct_now"),
                "foreign_streak_days": r.get("foreign_streak_days"),
                "foreign_streak_direction": r.get("foreign_streak_direction"),
                "latest_change_pct": r.get("latest_change_pct"),
            },
            description=f"외인 {r.get('foreign_streak_days', 0)}일 {r.get('foreign_streak_direction', '')}",
            icon="🇰🇷",
        )
        sao.priority = classify_priority(signals)
        objs.append(sao)
    print(f"  · investor_flow: {len(objs)}개 SAO (강력 시그널만)")
    return objs


def from_irr_analysis() -> list:
    """irr_analysis.json → KR_EQUITY (IRR 매수 후보) SAOs."""
    data = _load_json("irr_analysis.json")
    if not data:
        return []
    objs = []
    for r in data.get("results", []):
        mos = r.get("margin_of_safety_pct")
        if mos is None:
            continue
        # 안전마진 ≥ 5%pt → P2, ≥ 2%pt → P3
        signals = []
        if mos >= 5:
            signals.append(make_signal(
                "irr_strong_buy", 4, "🎯",
                "IRR 강력 매수 후보",
                f"안전마진 +{mos}%pt (IRR {r.get('irr', {}).get('primary_irr_pct')}%)"
            ))
        elif mos >= 2:
            signals.append(make_signal(
                "irr_watch", 3, "🟡",
                "IRR 매수 검토",
                f"안전마진 +{mos}%pt"
            ))
        else:
            continue  # MoS < 2 무시

        sao = AssetObject(
            id=stock_id(r["ticker"]),
            type="Stock",
            name=r["name"],
            pulse="KR_EQUITY",
            market=r.get("market"),
            sector=r.get("sector"),
            source=["yfinance"],
            last_updated=data.get("generated_at"),
            value=r.get("current_price"),
            unit="원",
            signals=signals,
            metrics={
                "primary_irr_pct": r.get("irr", {}).get("primary_irr_pct"),
                "required_return_pct": r.get("required_return", {}).get("required_return_pct"),
                "margin_of_safety_pct": mos,
                "rating": r.get("rating"),
            },
            icon="💎",
        )
        sao.priority = classify_priority(signals)
        objs.append(sao)
    print(f"  · irr_analysis: {len(objs)}개 SAO (MoS ≥ 2%pt)")
    return objs


def from_semi_challengers() -> list:
    """semi_challengers.json → SEMI Pulse SAOs."""
    data = _load_json("semi_challengers.json")
    if not data:
        return []
    objs = []
    for c in data.get("challengers", []):
        # 최근 뉴스가 있는 챌린저만 (활성 신호)
        news = c.get("recent_news", [])
        if not news:
            continue
        # 영향이 큰 챌린저 (negative+large 또는 magnitude=large 있음)
        max_impact = 0
        for imp in c.get("impacts", []):
            if imp.get("sentiment") == "negative" and imp.get("magnitude") == "large":
                max_impact = 4
                break
            elif imp.get("magnitude") == "large":
                max_impact = max(max_impact, 3)
            elif imp.get("magnitude") == "moderate":
                max_impact = max(max_impact, 2)

        signals = []
        if max_impact >= 4:
            signals.append(make_signal(
                "challenger_threat", 4, "⚔️",
                f"{c['name']} 위협 강함",
                f"{len(news)}건 신규 뉴스 + 큰 부정 영향"
            ))
        elif max_impact >= 3 or len(news) >= 2:
            signals.append(make_signal(
                "challenger_watch", 3, "🟡",
                f"{c['name']} 관찰",
                f"{len(news)}건 신규 뉴스"
            ))
        else:
            continue

        sao = AssetObject(
            id=f"challenger:{c['id']}",
            type="Company",
            name=c["name"],
            pulse="SEMI",
            sector=c.get("category"),
            source=[n.get("source", "RSS") for n in news[:3]],
            last_updated=data.get("generated_at"),
            value=(c.get("market") or {}).get("current_price"),
            unit="USD",
            signals=signals,
            metrics={
                "category": c.get("category"),
                "tech_label": c.get("tech_label"),
                "funding_total_usd_m": c.get("funding_total_usd_m"),
                "valuation_usd_b": c.get("valuation_usd_b"),
                "news_count": len(news),
                "impacts": c.get("impacts", []),
            },
            description=c.get("tech_description", ""),
            icon="🔬",
        )
        sao.priority = classify_priority(signals)
        objs.append(sao)
    print(f"  · semi_challengers: {len(objs)}개 SAO (활성 위협만)")
    return objs


def from_clarity_act() -> list:
    """clarity_act.json → LEGISLATIVE Pulse SAOs (법안만, 뉴스는 별도)."""
    data = _load_json("clarity_act.json")
    if not data:
        return []
    objs = []
    # 법안
    for b in data.get("bills", []):
        is_major = b.get("is_major_event")
        signals = []
        if is_major and b.get("status_code") == "enacted_signed":
            signals.append(make_signal(
                "bill_enacted", 5, "✅",
                f"{b['short_name']} 법률 제정",
                b.get("status_label", "")
            ))
        elif is_major:
            signals.append(make_signal(
                "bill_major_event", 4, b.get("status_emoji", "📋"),
                f"{b['short_name']} {b.get('status_label_kr', '')}",
                b.get("status_label", "")
            ))
        elif b.get("is_alive"):
            signals.append(make_signal(
                "bill_active", 3, "🟡",
                f"{b['short_name']} 추적 중",
                b.get("status_label", "")
            ))
        else:
            continue

        sao = AssetObject(
            id=bill_id(b["short_name"]),
            type="Bill",
            name=b["short_name"],
            pulse="LEGISLATIVE",
            source=["Congress.gov"],
            last_updated=data.get("generated_at"),
            signals=signals,
            metrics={
                "full_name": b.get("full_name"),
                "congress": b.get("congress"),
                "bill_type": b.get("bill_type"),
                "number": b.get("number"),
                "status_label_kr": b.get("status_label_kr"),
                "status_date": b.get("status_date"),
                "sponsor": b.get("sponsor"),
            },
            description=b.get("full_name", ""),
            link=b.get("link"),
            icon="🏛️",
        )
        sao.priority = classify_priority(signals)
        objs.append(sao)

    # 최근 뉴스 (상위 5건만 — Event Feed 노이즈 방지)
    for n in data.get("recent_news", [])[:5]:
        guid = n.get("guid") or n.get("link") or ""
        if not guid:
            continue
        signals = [make_signal(
            "legislation_news", 3, "📰",
            (n.get("title_ko") or n.get("title") or "")[:50],
            (n.get("summary_ko") or n.get("summary") or "")[:200],
        )]
        sao = AssetObject(
            id=news_id(guid),
            type="NewsItem",
            name=(n.get("title_ko") or n.get("title") or "")[:60],
            pulse="LEGISLATIVE",
            source=[n.get("source", "RSS")],
            last_updated=data.get("generated_at"),
            signals=signals,
            metrics={
                "title": n.get("title"),
                "title_ko": n.get("title_ko"),
                "summary_ko": n.get("summary_ko"),
                "matched_keyword": n.get("matched_keyword"),
            },
            link=n.get("link"),
            icon="📰",
        )
        sao.priority = classify_priority(signals)
        objs.append(sao)
    print(f"  · clarity_act: {len(objs)}개 SAO (법안 + 뉴스 상위 5건)")
    return objs


def from_bitcoin_standard() -> list:
    """bitcoin_standard.json → CRYPTO Pulse SAO."""
    data = _load_json("bitcoin_standard.json")
    if not data:
        return []
    bts = data.get("bts_score") or {}
    if not bts:
        return []
    severity = bts.get("severity")
    score = bts.get("score", 0)

    signals = []
    if severity == "red":
        signals.append(make_signal(
            "bitcoin_standard_alert", 4, "🚨",
            "비트코인 본위제 전환 신호 강함",
            f"BTS 점수 {score}/100 (Level: {bts.get('level')})"
        ))
    elif severity == "amber":
        signals.append(make_signal(
            "bitcoin_standard_watch", 3, "🟡",
            "비트코인 본위제 관찰",
            f"BTS 점수 {score}/100"
        ))
    else:
        return []  # green은 무시

    btc = data.get("btc") or {}
    sao = AssetObject(
        id="macro:BitcoinStandard",
        type="MacroEvent",
        name="비트코인 본위제 모니터",
        pulse="CRYPTO",
        source=["yfinance", "각종 거시"],
        last_updated=data.get("generated_at"),
        value=score,
        unit="/100",
        signals=signals,
        metrics={
            "btc_current": btc.get("current"),
            "btc_change_1d": btc.get("change_1d"),
            "btc_change_1y": btc.get("change_1y"),
            "level": bts.get("level"),
            "severity": severity,
        },
        icon="🪙",
    )
    sao.priority = classify_priority(signals)
    print(f"  · bitcoin_standard: 1개 SAO (severity={severity})")
    return [sao]


def from_breaking_news() -> list:
    """breaking_news_state.json → NEWS Pulse SAOs (최근 알린 뉴스만)."""
    state = _load_json("breaking_news_state.json")
    if not state:
        return []
    # 최근 알림 (last_alerts 같은 키가 있다고 가정)
    recent = state.get("last_alerts", []) or state.get("recent_news", []) or []
    # 5건만
    objs = []
    for n in recent[:5]:
        if isinstance(n, str):  # GUID만 저장된 경우
            continue
        if not isinstance(n, dict):
            continue
        guid = n.get("guid", "") or n.get("link", "") or n.get("title", "")[:50]
        signals = [make_signal(
            "breaking_news", 3, "🚨",
            (n.get("title") or "")[:50],
            n.get("matched_category", "긴급 뉴스"),
        )]
        sao = AssetObject(
            id=news_id(guid),
            type="NewsItem",
            name=(n.get("title") or "")[:60],
            pulse="NEWS",
            source=[n.get("source", "RSS")],
            last_updated=state.get("last_run"),
            signals=signals,
            link=n.get("link"),
            icon="📰",
        )
        sao.priority = classify_priority(signals)
        objs.append(sao)
    if objs:
        print(f"  · breaking_news: {len(objs)}개 SAO")
    return objs


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 72)
    print("  SAO Aggregator — Multi-Asset Fusion Center")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 72)

    print("\n[모듈별 SAO 추출]")
    all_objects = []
    all_objects.extend(from_credit_spread())
    all_objects.extend(from_investor_flow())
    all_objects.extend(from_irr_analysis())
    all_objects.extend(from_semi_challengers())
    all_objects.extend(from_clarity_act())
    all_objects.extend(from_bitcoin_standard())
    all_objects.extend(from_breaking_news())

    print(f"\n  총 SAO: {len(all_objects)}개")

    # 페이로드 빌드
    payload = build_signals_payload(all_objects)

    # 저장
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    # 요약
    print(f"\n[Event Feed 요약]")
    print(f"  Ready to Act (P1+P2): {payload['ready_to_act']}건")
    print(f"  우선순위별:")
    for p, c in payload["priority_counts"].items():
        print(f"    {p}: {c}건")
    print(f"  Pulse별:")
    for p, c in payload["pulse_counts"].items():
        if c > 0:
            print(f"    {p}: {c}건")

    # Top 5 미리보기
    print(f"\n[Top 5 시그널 미리보기]")
    for s in payload["signals"][:5]:
        first_sig = s["signals"][0] if s["signals"] else {}
        print(f"  {s['priority']} [{s['pulse']:10s}] {s['name'][:30]:30s} — {first_sig.get('label', '')}")

    print("\n" + "=" * 72)
    print(f"  완료: SAO {len(all_objects)}개 통합")
    print("=" * 72)


if __name__ == "__main__":
    main()
