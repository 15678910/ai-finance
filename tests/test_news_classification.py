"""뉴스 분류·선별 로직 테스트.

대상:
  · news_impact.py   — 격화/완화 판정 (단어 경계 매칭)
  · market_news.py   — 카테고리 분류, 노이즈 판정, 중요도 가중

실행: python -m pytest tests/test_news_classification.py -v
또는: python tests/test_news_classification.py
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


# ====================================================================
# 1. 부분문자열 오분류 (2026-08 실측 버그)
# ====================================================================
def test_substring_false_positives_fixed():
    """키워드가 다른 단어 안에 들어 있어도 매칭되던 버그.

    전부 실제로 잘못 분류되던 사례다. 예: 'increase' 안의 'ease' 때문에
    금리 인상이 완화(호재)로 분류됐다.
    """
    from news_impact import classify_news

    cases = [
        ("Fed increases rates by 25bp", "increase ⊃ ease"),
        ("Microsoft software update boosts productivity", "software ⊃ war"),
        ("An ideal entry point for long-term investors", "ideal ⊃ deal"),
        ("Oil prices route through new pipeline", "route ⊃ rout"),
        ("Disease outbreak halts factory output", "disease ⊃ ease"),
        ("Please note: earnings released today", "please ⊃ ease"),
        ("Warehouse automation lifts margins", "warehouse ⊃ war"),
    ]
    for title, why in cases:
        score, label, _, _ = classify_news(title)
        assert score == 0, f"{title!r} → {label} (원인: {why})"


def test_crowdstrike_not_geopolitical():
    """'CrowdStrike' 안의 strike 가 지정학으로 잡히던 버그."""
    from market_news import _categorize

    cat, _ = _categorize("Dow Jones Futures: Market Dips As CrowdStrike Tumbles")
    assert cat != "지정학"


def test_fedex_not_fed_category():
    """'FedEx' 안의 fed 가 Fed·물가로 잡히던 버그."""
    from market_news import _categorize

    cat, _ = _categorize("FedEx delivery volumes rise in Q3")
    assert cat != "Fed·물가"


# ====================================================================
# 2. 정상 판정 회귀 방지
# ====================================================================
def test_genuine_signals_still_detected():
    from news_impact import classify_news

    escalation = [
        "Israel strikes Iran nuclear site",
        "Market plunges on escalating tensions",
        "Trump warns of severe economic actions",
        "US to impose toughest sanctions in history",
        "Nasdaq tumbles as chip stocks slump",
        "이란 미사일 공격에 증시 급락",
    ]
    for title in escalation:
        score, label, _, _ = classify_news(title)
        assert score < 0, f"{title!r} → {label} (격화여야 함)"

    deescalation = [
        "US and China reach trade deal",
        "Ceasefire agreement holds in Gaza",
        "Stocks rally as inflation eases",
        "휴전 합의 기대감에 반등",
        "금리 인하 기대가 완화됐다",       # 한글은 조사·어미가 붙어도 잡혀야 한다
    ]
    for title in deescalation:
        score, label, _, _ = classify_news(title)
        assert score > 0, f"{title!r} → {label} (완화여야 함)"


def test_negation_patterns():
    """'완화 무산'=악재 / '공격 취소'=호재."""
    from news_impact import classify_news

    assert classify_news("Deal hopes fade as talks collapse")[0] < 0
    assert classify_news("Israeli strikes canceled after agreement")[0] > 0


def test_mixed_signal_is_neutral():
    from news_impact import classify_news

    score, label, _, _ = classify_news("Peace talks resume after missile attack")
    assert score == 0 and label == "혼조"


def test_stem_keywords_match_inflections():
    """어간 표기(escalat*)는 활용형을 모두 잡아야 한다."""
    from news_impact import classify_news

    for title in ["Tensions escalate", "Escalating conflict", "Further escalation"]:
        assert classify_news(title)[0] < 0, title


# ====================================================================
# 3. 카테고리 커버리지 (수집 소스가 없어 항상 0건이던 주제)
# ====================================================================
def test_fed_and_crypto_categories_reachable():
    """베센트 국채 발언·트럼프 암호화폐 발언이 제 카테고리로 분류되는지."""
    from market_news import _categorize

    assert _categorize("Bessent says Treasury yields will stabilize")[0] == "Fed·물가"
    assert _categorize("베센트 재무장관, 국채 발행 계획 발표")[0] == "Fed·물가"
    assert _categorize("Powell signals rate cut at Jackson Hole")[0] == "Fed·물가"
    assert _categorize("Trump touts bitcoin reserve plan")[0] == "암호화폐"
    assert _categorize("SEC approves spot ETF for digital assets")[0] == "암호화폐"


def test_topic_feeds_cover_fed_and_crypto():
    """수집 피드에 연준·국채·암호화폐 쿼리가 실제로 들어 있는지.

    분류기에 카테고리가 있어도 그 뉴스를 가져오는 피드가 없으면 영원히 0건이다.
    """
    from market_news import TOPIC_FEEDS

    joined = " ".join(url for url, _ in TOPIC_FEEDS).lower()
    for term in ("bessent", "powell", "treasury", "fomc", "bitcoin", "crypto"):
        assert term in joined, f"'{term}' 를 찾는 피드가 없다"


# ====================================================================
# 4. 노이즈 판정 · 중요도 가중
# ====================================================================
def test_noise_detection():
    """키워드만 우연히 걸린 지역·스포츠 기사 (실측 사례)."""
    from market_news import is_noise

    assert is_noise("The Epicenter: Women's Global Sports Summit - KPTV")
    assert is_noise("Apparent lightning strike ignites attic fire at home in Summit, NE")
    assert not is_noise("US says it will impose toughest sanctions on Iran")
    assert not is_noise("Bessent comments on Treasury yields")


def test_priority_ranks_market_movers_above_columns():
    """시장을 움직인 헤드라인이 투자 칼럼보다 위로 와야 한다."""
    from market_news import priority_score

    fed_mover = priority_score("Fed·물가", -1, False)      # 방향성 있는 정책 뉴스
    geo_mover = priority_score("지정학", -1, False)
    column = priority_score("기타", 0, False)              # "If a Bear Market Is Coming…"
    noise = priority_score("지정학", -1, True)             # 낙뢰 strike 기사

    assert fed_mover > geo_mover > column > noise
    assert noise < 0, "노이즈는 음수라야 24건 상한에서 밀려난다"


def test_priority_directional_bonus():
    """같은 카테고리면 방향성 판정이 붙은 쪽이 우선."""
    from market_news import priority_score

    assert priority_score("지정학", -1, False) > priority_score("지정학", 0, False)


# ====================================================================
# 5. 순심리 집계
# ====================================================================
def test_aggregate_sentiment():
    from news_impact import aggregate_sentiment

    esc = aggregate_sentiment(["Israel strikes Iran", "Market plunges", "War escalates"])
    assert esc["score"] <= -2 and "위험" in esc["label"]

    de = aggregate_sentiment(["Ceasefire agreement", "Stocks rally", "Trade deal reached"])
    assert de["score"] >= 2 and "우호" in de["label"]

    assert aggregate_sentiment([])["label"] == "중립"


def test_sentiment_keyword_symmetry():
    """상승/하락 시세 동사가 대칭이라야 심리가 한쪽으로 치우치지 않는다."""
    from news_impact import classify_news

    assert classify_news("Stocks surge on optimism")[0] > 0
    assert classify_news("Stocks tumble on fears")[0] < 0
    assert classify_news("Market rallies")[0] > 0
    assert classify_news("Market slumps")[0] < 0


if __name__ == "__main__":
    import traceback

    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    passed = 0
    for test in tests:
        try:
            test()
            print(f"  ✓ {test.__name__}")
            passed += 1
        except Exception:
            print(f"  ✗ {test.__name__}")
            traceback.print_exc()

    print(f"\n{passed}/{len(tests)} 통과")
    sys.exit(0 if passed == len(tests) else 1)
