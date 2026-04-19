"""
AI 투자 코멘트 엔진 (규칙 기반)
================================
매크로 + 섹터 + 지정학 + 포트폴리오 데이터를 교차 분석하여
투자 판단 코멘트를 자동 생성합니다.
API 없이 무료로 동작합니다.
"""

import json
from datetime import datetime


def generate_commentary(data: dict) -> dict:
    """data.json 구조를 받아 AI 코멘트를 생성합니다."""

    macro = data.get("macro", {})
    macro_detail = data.get("macro_detail", {})
    geo = data.get("geopolitical", {})
    sectors = data.get("sectors", [])
    portfolios = data.get("portfolios", {})
    insights = data.get("insights", {})

    macro_regime = _classify_macro_regime(macro, macro_detail)

    commentary = {
        "macro_regime": macro_regime,
        "market_overview": _generate_market_overview(macro, macro_detail, geo),
        "sector_picks": _generate_sector_picks(sectors, macro, geo, portfolios, macro_regime),
        "risk_assessment": _generate_risk_assessment(macro_detail, geo),
        "action_items": _generate_action_items(macro, macro_detail, geo, sectors, portfolios, macro_regime),
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    return commentary


def _classify_macro_regime(macro, macro_detail) -> dict:
    """매크로 체제 분류: 금융 억압 vs 정상 시장 (러셀 내피어 프레임워크)"""
    # 실질금리 파싱
    real_rate_val = None
    real_rate_str = macro_detail.get("real_rate", "") if macro_detail else ""
    try:
        real_rate_val = float(str(real_rate_str).replace("%", "").strip())
    except (ValueError, TypeError):
        pass

    # 10년 국채 금리
    treasury_10y = None
    t10_data = macro_detail.get("treasury_10y", {}) if macro_detail else {}
    if isinstance(t10_data, dict):
        treasury_10y = t10_data.get("current")
    elif isinstance(t10_data, (int, float)):
        treasury_10y = t10_data

    # CPI
    cpi_val = None
    cpi_str = macro_detail.get("cpi_yoy", "") if macro_detail else ""
    try:
        cpi_val = float(str(cpi_str).replace("%", "").strip())
    except (ValueError, TypeError):
        pass

    # 달러 인덱스
    dollar_idx = None
    dx = macro_detail.get("dollar_index", {}) if macro_detail else {}
    if isinstance(dx, dict):
        dollar_idx = dx.get("current")
    elif isinstance(dx, (int, float)):
        dollar_idx = dx

    # 체제 판정
    regime = "정상 시장"
    severity = "green"
    description = ""
    beneficiaries = []
    losers = []

    if real_rate_val is not None:
        if real_rate_val < 0:
            regime = "금융 억압 국면"
            severity = "red"
            description = f"실질금리 {real_rate_val:+.2f}%로 마이너스. 정부가 인플레이션보다 낮은 금리로 부채 실질 가치를 희석하는 국면입니다. 저축자(채권보유자)의 부가 채무자(정부)로 이전됩니다."
            beneficiaries = ["방산", "재산업화(제조업/반도체)", "에너지/원자재", "금/귀금속", "실물자산", "가치주"]
            losers = ["장기 국채", "고평가 성장주", "현금성 자산"]
        elif real_rate_val < 1.0:
            regime = "약한 금융 억압"
            severity = "amber"
            description = f"실질금리 {real_rate_val:+.2f}%로 낮은 수준. 명시적 억압은 아니나 저축자에게 불리한 환경이 지속되고 있습니다."
            beneficiaries = ["방산", "에너지", "금", "인프라", "배당 가치주"]
            losers = ["장기채권", "저배당 성장주"]
        else:
            regime = "정상 시장"
            severity = "green"
            description = f"실질금리 {real_rate_val:+.2f}%로 양(+). 자유 시장 가격 결정 메커니즘이 정상 작동 중입니다."
            beneficiaries = ["성장주", "장기채권", "부동산"]
            losers = ["원자재 일부", "금"]

    return {
        "regime": regime,
        "severity": severity,
        "description": description,
        "real_rate": real_rate_val,
        "treasury_10y": treasury_10y,
        "cpi": cpi_val,
        "dollar_index": dollar_idx,
        "beneficiaries": beneficiaries,
        "losers": losers,
    }


def _generate_market_overview(macro, macro_detail, geo) -> str:
    """시장 종합 평가 코멘트"""
    cycle = macro.get("cycle", "N/A")
    rate = macro.get("rate", "N/A")
    cpi = macro.get("cpi", "N/A")
    geo_score = geo.get("risk_score", 0) or 0
    geo_level = geo.get("risk_level", "N/A")

    vix = None
    if macro_detail:
        vix_data = macro_detail.get("vix", {})
        if isinstance(vix_data, dict):
            vix = vix_data.get("current")
        elif isinstance(vix_data, (int, float)):
            vix = vix_data

    parts = []

    # Economic cycle assessment
    if "확장" in str(cycle):
        parts.append("경기 확장기로 성장 모멘텀이 유효합니다")
    elif "수축" in str(cycle) or "정점" in str(cycle):
        parts.append("경기 둔화 신호가 감지되어 방어적 포지션이 필요합니다")
    else:
        parts.append(f"현재 경기 사이클은 {cycle} 국면입니다")

    # Rate cycle impact
    if "인하" in str(rate):
        parts.append("금리 인하기 진입으로 성장주와 장기채에 우호적인 환경입니다")
    elif "인상" in str(rate):
        parts.append("금리 인상 압력으로 밸류에이션 부담이 존재합니다")
    else:
        parts.append("금리 동결 국면으로 시장은 관망세를 보이고 있습니다")

    # Geopolitical risk
    if geo_score >= 70:
        parts.append(f"지정학 리스크가 {geo_level}({geo_score:.0f}/100) 수준으로 안전자산 선호가 강화될 수 있습니다")
    elif geo_score >= 50:
        parts.append(f"지정학 리스크 {geo_level}({geo_score:.0f}/100)으로 변동성 확대에 유의해야 합니다")
    elif geo_score > 0:
        parts.append(f"지정학 리스크는 {geo_level}({geo_score:.0f}/100)으로 시장에 미치는 영향은 제한적입니다")

    # VIX
    if vix is not None:
        if vix > 30:
            parts.append(f"VIX {vix:.1f}로 공포 수준이며 단기 변동성이 매우 높습니다")
        elif vix > 25:
            parts.append(f"VIX {vix:.1f}로 경계 수준이며 헤지 전략을 고려해야 합니다")
        elif vix > 20:
            parts.append(f"VIX {vix:.1f}로 불안정한 시장 심리를 반영합니다")

    return ". ".join(parts) + "."


def _generate_sector_picks(sectors, macro, geo, portfolios, macro_regime=None) -> list:
    """섹터별 투자 코멘트"""
    picks = []
    rate = str(macro.get("rate", ""))
    cycle = str(macro.get("cycle", ""))
    geo_score = geo.get("risk_score", 0) or 0

    # Get portfolio sharpe data
    pf_sectors = {}
    if isinstance(portfolios, dict) and "sectors" in portfolios:
        for s in portfolios.get("sectors", []):
            pf_sectors[s.get("name", "")] = s.get("sharpe_ratio", 0)

    for sector in sectors:
        name = sector.get("name", "")
        stocks = sector.get("stocks", [])
        if not stocks:
            continue

        # Calculate sector averages
        sentiment_scores = []
        regimes = []
        for stock in stocks:
            score = stock.get("sentiment_score")
            if score is not None:
                try:
                    sentiment_scores.append(float(str(score).replace('+', '')))
                except (ValueError, TypeError):
                    pass
            regime = stock.get("regime", "")
            if regime:
                regimes.append(regime)

        avg_sentiment = sum(sentiment_scores) / len(sentiment_scores) if sentiment_scores else 0

        # Count bullish/bearish regimes
        bullish = sum(1 for r in regimes if "상승" in r)
        bearish = sum(1 for r in regimes if "하락" in r)

        # Find sharpe ratio
        sharpe = None
        for pf_name, pf_sharpe in pf_sectors.items():
            if any(part in pf_name for part in name.replace("/", "_").split("_")):
                sharpe = pf_sharpe
                break

        comment_parts = []
        signal = "중립"

        # Regime-based
        if bullish > bearish:
            comment_parts.append("상승 레짐 우세로 모멘텀 유리")
            signal = "긍정"
        elif bearish > bullish:
            comment_parts.append("하락 레짐 우세로 신중한 접근 필요")
            signal = "부정"
        else:
            comment_parts.append("횡보 국면으로 방향성 탐색 중")

        # Sentiment-based
        if avg_sentiment > 20:
            comment_parts.append("시장 심리 낙관적")
        elif avg_sentiment < -20:
            comment_parts.append("시장 심리 비관적 — 역발상 매수 기회 탐색")

        # Macro cross-reference
        if "인하" in rate:
            if "IT" in name or "반도체" in name or "빅테크" in name:
                comment_parts.append("금리 인하기 기술주 밸류에이션 상승 기대")
                signal = "긍정"
            elif "채권" in name or "배당" in name:
                comment_parts.append("금리 인하 시 채권/배당주 매력 상승")
        elif "인상" in rate:
            if "에너지" in name or "은행" in name:
                comment_parts.append("금리 인상기 에너지/금융 섹터 수혜 가능")

        # Geopolitical cross-reference
        if geo_score >= 60:
            if "방산" in name:
                comment_parts.append("지정학 리스크 상승으로 방산 수혜 예상")
                signal = "긍정"
            elif "에너지" in name:
                comment_parts.append("지정학 긴장으로 에너지 가격 변동성 확대")
            elif "암호화폐" in name:
                comment_parts.append("지정학 불안 시 위험자산 회피 심리 주의")
                signal = "부정" if signal != "긍정" else signal

        # Sharpe ratio
        if sharpe is not None:
            if sharpe > 1.5:
                comment_parts.append(f"Sharpe {sharpe:.2f}로 위험 대비 수익 우수")
            elif sharpe < 0:
                comment_parts.append(f"Sharpe {sharpe:.2f}로 위험 대비 수익 열악")

        # 금융 억압 체제 기반 코멘트 (러셀 내피어 프레임워크)
        if macro_regime and macro_regime.get("regime") in ("금융 억압 국면", "약한 금융 억압"):
            is_benef = False
            is_loser = False
            if "방산" in name:
                is_benef = True
                comment_parts.append("금융 억압 국면 재무장 수혜 명확")
            elif "에너지" in name or "원자재" in name:
                is_benef = True
                comment_parts.append("실물 자산으로 인플레이션 헤지 효과")
            elif "반도체" in name or "IT" in name or "배터리" in name or "2차전지" in name:
                is_benef = True
                comment_parts.append("재산업화(Re-industrialization) 수혜 대상")
            elif "암호화폐" in name:
                comment_parts.append("자본 통제 회피 수단으로 분산 고려")

            if is_benef and signal == "중립":
                signal = "긍정"
            if is_benef and macro_regime.get("severity") == "red":
                comment_parts.append("실질금리 마이너스 환경에서 비중 확대 권고")

        # Top stock picks within sector
        top_stocks = []
        for stock in stocks:
            s_name = stock.get("name", "")
            s_regime = stock.get("regime", "")
            s_sent = stock.get("sentiment_label", "")
            if "상승" in s_regime and ("긍정" in str(s_sent) or "낙관" in str(s_sent)):
                top_stocks.append(s_name)

        pick = {
            "sector": name,
            "signal": signal,
            "comment": ". ".join(comment_parts) + ".",
        }
        if top_stocks:
            pick["top_picks"] = top_stocks[:2]

        picks.append(pick)

    return picks


def _generate_risk_assessment(macro_detail, geo) -> list:
    """리스크 평가 코멘트"""
    risks = []

    geo_score = geo.get("risk_score", 0) or 0
    categories = geo.get("categories", [])

    # Geopolitical risks
    for cat in categories:
        cat_name = cat.get("category") or cat.get("name", "")
        cat_score = cat.get("score", 0) or 0
        if cat_score >= 30:
            risks.append({
                "type": "지정학",
                "severity": "높음" if cat_score >= 50 else "중간",
                "description": f"{cat_name} 리스크 점수 {cat_score}으로 관련 섹터 영향 주의"
            })

    # VIX risk
    if macro_detail:
        vix_data = macro_detail.get("vix", {})
        vix = vix_data.get("current") if isinstance(vix_data, dict) else vix_data if isinstance(vix_data, (int, float)) else None
        if vix and vix > 25:
            risks.append({
                "type": "시장변동성",
                "severity": "높음" if vix > 30 else "중간",
                "description": f"VIX {vix:.1f}로 {'공포' if vix > 30 else '경계'} 수준. 포지션 사이즈 축소 권고"
            })

    # Yield curve risk
    if macro_detail:
        yc = macro_detail.get("yield_curve", {})
        if isinstance(yc, dict):
            spread = yc.get("spread_10y2y")
            if spread is not None and spread < 0:
                risks.append({
                    "type": "경기침체",
                    "severity": "높음",
                    "description": f"장단기 금리 역전(스프레드 {spread:.2f}%p) — 경기 침체 경고 신호"
                })

    # Inflation risk
    if macro_detail:
        cpi = macro_detail.get("cpi_yoy", "")
        try:
            cpi_val = float(str(cpi).replace("%", ""))
            if cpi_val > 4:
                risks.append({
                    "type": "인플레이션",
                    "severity": "높음",
                    "description": f"CPI {cpi_val:.1f}%로 고인플레이션. 실질 수익률 하락 주의"
                })
        except (ValueError, TypeError):
            pass

    return risks


def _generate_action_items(macro, macro_detail, geo, sectors, portfolios, macro_regime=None) -> list:
    """구체적 투자 행동 권고"""
    actions = []

    rate = str(macro.get("rate", ""))
    cycle = str(macro.get("cycle", ""))
    geo_score = geo.get("risk_score", 0) or 0

    # Rate cycle actions
    if "인하" in rate:
        actions.append({
            "priority": "높음",
            "action": "금리 인하 수혜주 비중 확대",
            "detail": "성장주(기술/AI), 장기채(10Y+), 금/귀금속 비중을 점진적으로 확대하세요."
        })
        actions.append({
            "priority": "중간",
            "action": "달러 약세 대비 포지션 조정",
            "detail": "원화 강세 가능성에 따라 해외자산 환헤지 비율을 재검토하세요."
        })
    elif "인상" in rate:
        actions.append({
            "priority": "높음",
            "action": "금리 인상 방어 포지션",
            "detail": "단기채 비중 확대, 고PER 성장주 비중 축소를 검토하세요."
        })

    # 금융 억압 체제 기반 행동 권고 (러셀 내피어 프레임워크)
    if macro_regime:
        regime_name = macro_regime.get("regime", "")
        real_rate = macro_regime.get("real_rate")
        if regime_name == "금융 억압 국면":
            actions.append({
                "priority": "높음",
                "action": "장기 채권 비중 축소 (금융 억압 국면)",
                "detail": f"실질금리 {real_rate:+.2f}%로 마이너스. 장기 국채 실질 가치 손실 경고. 밸류 주식과 실물 자산으로 대체하세요."
            })
            actions.append({
                "priority": "높음",
                "action": "재산업화/재무장 수혜주 비중 확대",
                "detail": "국가 자본주의 체제에서 정부가 자본을 집중하는 방산, 반도체, 배터리, 에너지 인프라 섹터에 투자하세요."
            })
            actions.append({
                "priority": "중간",
                "action": "금/귀금속 배분 검토 (인플레이션 헤지)",
                "detail": "금융 억압 + 자본 통제 방어 수단. 포트폴리오의 5-10%를 금으로 배분을 고려하세요."
            })
        elif regime_name == "약한 금융 억압":
            actions.append({
                "priority": "중간",
                "action": "채권 비중 점진적 축소",
                "detail": f"실질금리 {real_rate:+.2f}%로 낮은 수준. 장기채 듀레이션을 줄이고 가치주/배당주 비중을 늘리세요."
            })

    # Geopolitical actions
    if geo_score >= 70:
        actions.append({
            "priority": "높음",
            "action": "안전자산 비중 확대 권고",
            "detail": f"지정학 리스크 {geo_score:.0f}/100 심각 수준. 금, 국채, 방산주 비중을 늘리고 위험자산 노출을 줄이세요."
        })
    elif geo_score >= 50:
        actions.append({
            "priority": "중간",
            "action": "포트폴리오 방어력 점검",
            "detail": f"지정학 리스크 {geo_score:.0f}/100 경계 수준. 스톱로스 설정 및 현금 비중 10-15% 유지를 권고합니다."
        })

    # Sector-specific actions based on extreme sentiment
    for sector in sectors:
        name = sector.get("name", "")
        stocks = sector.get("stocks", [])
        for stock in stocks:
            sent_label = str(stock.get("sentiment_label", ""))
            sent_score = stock.get("sentiment_score")
            stock_name = stock.get("name", "")

            try:
                score_val = float(str(sent_score).replace('+', '')) if sent_score else 0
            except (ValueError, TypeError):
                score_val = 0

            if "극도" in sent_label and "공포" in sent_label:
                actions.append({
                    "priority": "중간",
                    "action": f"{stock_name} 역발상 매수 기회 탐색",
                    "detail": f"극도 공포({score_val:+.1f}) 구간은 역사적으로 저점 매수 기회. 분할 매수 검토."
                })
            elif "극도" in sent_label and "낙관" in sent_label:
                actions.append({
                    "priority": "중간",
                    "action": f"{stock_name} 과열 주의",
                    "detail": f"극도 낙관({score_val:+.1f}) 구간은 과열 신호. 부분 차익 실현 검토."
                })

    # Portfolio rebalancing
    if isinstance(portfolios, dict) and "best_sector" in portfolios and "worst_sector" in portfolios:
        best = portfolios.get("best_sector", {})
        worst = portfolios.get("worst_sector", {})
        if best and worst:
            actions.append({
                "priority": "낮음",
                "action": "섹터 리밸런싱 검토",
                "detail": f"최고 Sharpe: {best.get('name', '')}({best.get('sharpe_ratio', 0):.2f}), 최저: {worst.get('name', '')}({worst.get('sharpe_ratio', 0):.2f}). 저성과 섹터에서 고성과 섹터로 비중 이동을 고려하세요."
            })

    return actions


# CLI 테스트용
if __name__ == "__main__":
    import sys
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    path = sys.argv[1] if len(sys.argv) > 1 else "docs/data.json"
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    result = generate_commentary(data)
    print(json.dumps(result, ensure_ascii=False, indent=2))
