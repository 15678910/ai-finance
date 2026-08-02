"""가치 투자 전략 로직 테스트.

대상:
  · value_screener.py — 시장 상대 밸류에이션, 가치 함정 필터, 소외도
  · value_erosion_monitor.py — 가치 훼손 판정과 기회/경보 분류

실행: python -m pytest tests/test_value_strategy.py -v
또는: python tests/test_value_strategy.py
"""

import os
import sys
import json
import tempfile
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


def _write_baseline(per=18.0, pbr=1.9, asof="2026-08-01"):
    """임시 kospi_valuation.json 생성."""
    fd, path = tempfile.mkstemp(suffix=".json")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump({"per": per, "pbr": pbr, "asof": asof}, f)
    return path


# ====================================================================
# 1. 시장 상대 밸류에이션 기준선
# ====================================================================
def test_baseline_anchors():
    """앵커 = max(시장평균 -30%, 금리 2배 역수)."""
    import value_screener as vs

    path = _write_baseline(per=18.0, pbr=1.9)
    try:
        b = vs.load_market_baseline(path=path, today=datetime(2026, 8, 2))
        assert b["available"] is True
        assert b["per_market_anchor"] == 12.6           # 18 × 0.7
        assert b["per_rate_anchor"] == 14.29            # 1 / (2 × 3.5%)
        assert b["per_anchor"] == 14.29                 # 둘 중 느슨한 쪽
        assert b["pbr_anchor"] == 1.9                   # 시장 평균 그대로
    finally:
        os.unlink(path)


def test_baseline_rejects_stale_and_missing():
    """낡거나 없는 데이터는 available=False → 절대 임계값 폴백."""
    import value_screener as vs

    assert vs.load_market_baseline(path="/nonexistent.json")["available"] is False

    path = _write_baseline(asof="2026-01-01")
    try:
        b = vs.load_market_baseline(path=path, today=datetime(2026, 8, 2))
        assert b["available"] is False
        assert "경과" in b["reason"]
    finally:
        os.unlink(path)


def test_relative_scoring_tracks_market_level():
    """같은 PER이라도 시장 레벨에 따라 점수가 달라져야 한다.

    고정 임계값의 결함을 고친 것이 이 변경의 핵심이므로 명시적으로 검증한다.
    """
    import value_screener as vs

    cheap_market = _write_baseline(per=8.0, pbr=0.9)     # 시장 전체가 싼 장세
    rich_market = _write_baseline(per=25.0, pbr=2.5)     # 시장 전체가 비싼 장세
    try:
        today = datetime(2026, 8, 2)
        b_cheap = vs.load_market_baseline(path=cheap_market, today=today)
        b_rich = vs.load_market_baseline(path=rich_market, today=today)

        # PER 14배: 싼 장세에서는 평범, 비싼 장세에서는 저평가
        assert vs.score_per(14.0, b_cheap) < vs.score_per(14.0, b_rich)

        # PBR 1.5배도 마찬가지
        assert vs.score_pbr(1.5, b_cheap) < vs.score_pbr(1.5, b_rich)
    finally:
        os.unlink(cheap_market)
        os.unlink(rich_market)


def test_scoring_falls_back_to_absolute():
    """기준선이 없으면 기존 절대 임계값 그대로."""
    import value_screener as vs

    assert vs.score_per(4.0, None) == 100
    assert vs.score_per(35.0, None) == 5
    assert vs.score_pbr(0.4, None) == 100
    assert vs.score_pbr(5.0, None) == 10
    # 무효값
    assert vs.score_per(None) == 0
    assert vs.score_per(-3) == 0
    assert vs.score_pbr(0) == 0


def test_scoring_mode_reported():
    """총점 결과에 상대/절대 모드가 기록되어야 한다."""
    import value_screener as vs

    metrics = {"trailing_pe": 10, "price_to_book": 1.0, "roe": 0.15,
               "debt_to_equity": 50, "revenue_growth": 0.1, "dividend_yield": 2.0}
    assert vs.calculate_value_score(metrics)["scoring_mode"] == "absolute"

    path = _write_baseline()
    try:
        b = vs.load_market_baseline(path=path, today=datetime(2026, 8, 2))
        assert vs.calculate_value_score(metrics, b)["scoring_mode"] == "relative"
    finally:
        os.unlink(path)


# ====================================================================
# 2. 가치 함정 필터
# ====================================================================
def _metrics(**quality):
    return {"market_cap": 5e12, "roe": 0.10, "debt_to_equity": 80, "quality": quality}


def test_trap_hard_reject_on_chronic_loss():
    """3년 연속 적자는 배제. 2년은 감점 후 통과."""
    import value_screener as vs

    passed, reason, flags, penalty = vs.value_trap_filter(_metrics(consecutive_loss_years=3))
    assert passed is False
    assert "연속 적자" in reason

    passed, _, flags, penalty = vs.value_trap_filter(_metrics(consecutive_loss_years=2))
    assert passed is True
    assert penalty == 15
    assert flags[0]["type"] == "적자지속"


def test_trap_soft_flags_accumulate():
    """사양 산업·자산의 질·장부가 허수는 감점으로 누적."""
    import value_screener as vs

    passed, reason, flags, penalty = vs.value_trap_filter(_metrics(
        revenue_cagr_pct=-8.0, revenue_years=4,
        inventory_receivable_ratio_pct=60.0,
        intangible_to_equity_pct=70.0,
    ))
    assert passed is True
    assert penalty == 12 + 8 + 10
    assert {f["type"] for f in flags} == {"사양산업", "자산의질", "장부가허수"}
    assert "경고" in reason


def test_trap_clean_stock_passes_without_penalty():
    import value_screener as vs

    passed, reason, flags, penalty = vs.value_trap_filter(_metrics(
        consecutive_loss_years=0, revenue_cagr_pct=12.0,
        inventory_receivable_ratio_pct=20.0, intangible_to_equity_pct=5.0,
    ))
    assert (passed, reason, flags, penalty) == (True, "통과", [], 0)


def test_trap_filter_survives_missing_quality_data():
    """재무제표 수집 실패 시에도 기존 하드 룰은 동작해야 한다."""
    import value_screener as vs

    passed, _, flags, penalty = vs.value_trap_filter({"market_cap": 5e12, "roe": 0.1})
    assert (passed, flags, penalty) == (True, [], 0)

    passed, reason, _, _ = vs.value_trap_filter({"market_cap": 5e10})
    assert passed is False and "시가총액" in reason


def test_quality_data_extraction():
    """재무제표 → 함정 판정 원천 데이터. 계정명 변종·결측에 견뎌야 한다."""
    import pandas as pd
    import value_screener as vs

    class FakeTicker:
        def __init__(self, income, bs, raise_on=None):
            self._income, self._bs, self._raise = income, bs, raise_on

        @property
        def income_stmt(self):
            if self._raise == "income":
                raise RuntimeError("네트워크 오류")
            return self._income

        @property
        def balance_sheet(self):
            if self._raise == "bs":
                raise RuntimeError("네트워크 오류")
            return self._bs

    income = pd.DataFrame(
        {"2025": [-500, 8000], "2024": [-300, 9000],
         "2023": [200, 11000], "2022": [400, 12000]},
        index=["Net Income", "Total Revenue"])
    bs = pd.DataFrame(
        {"2025": [60000, 20000, 15000, 30000, 18000]},
        index=["Total Assets", "Inventory", "Accounts Receivable",
               "Stockholders Equity", "Goodwill"])

    q = vs.fetch_quality_data(FakeTicker(income, bs))
    assert q["consecutive_loss_years"] == 2          # 2025·2024 적자, 2023 흑자에서 중단
    assert q["revenue_cagr_pct"] < 0                 # 매출 12000 → 8000
    assert q["inventory_receivable_ratio_pct"] == 58.3
    assert q["intangible_to_equity_pct"] == 60.0
    assert q["data_ok"] is True

    # 수집 실패해도 예외 없이 기본값
    assert vs.fetch_quality_data(FakeTicker(None, None))["data_ok"] is False
    assert vs.fetch_quality_data(FakeTicker(income, bs, "bs"))["consecutive_loss_years"] == 2


def test_statement_row_handles_variants():
    """계정명 중복·NaN·문자열 등 실데이터 변종."""
    import pandas as pd
    import value_screener as vs

    df = pd.DataFrame({"2025": [100, 200], "2024": [90, 180]},
                      index=["Net Income", "Total Revenue"])
    assert vs._statement_row(df, "Net Income") == [100.0, 90.0]
    assert vs._statement_row(df, "없는계정", "Total Revenue") == [200.0, 180.0]
    assert vs._statement_row(df, "없음") == []
    assert vs._statement_row(None, "Net Income") == []

    dup = pd.DataFrame({"2025": [100, 999]}, index=["Net Income", "Net Income"])
    assert vs._statement_row(dup, "Net Income") == [100.0]

    nan_df = pd.DataFrame({"2025": [100.0], "2024": [float("nan")]}, index=["Net Income"])
    assert vs._statement_row(nan_df, "Net Income") == [100.0]

    bad = pd.DataFrame({"2025": ["N/A"]}, index=["Net Income"])
    assert vs._statement_row(bad, "Net Income") == []


# ====================================================================
# 3. 소외도
# ====================================================================
def test_neglect_score_ranks_forgotten_stocks_higher():
    import value_screener as vs

    forgotten = vs.calculate_neglect_score({
        "analyst_count": 0, "average_volume": 50_000,
        "current_price": 20_000, "market_cap": 5e11})
    crowded = vs.calculate_neglect_score({
        "analyst_count": 35, "average_volume": 12_000_000,
        "current_price": 80_000, "market_cap": 480e12})

    assert forgotten["score"] > crowded["score"]
    assert forgotten["partial"] is False


def test_neglect_partial_and_missing_data():
    """한 축만 있으면 그 축으로, 둘 다 없으면 available=False."""
    import value_screener as vs

    partial = vs.calculate_neglect_score({"analyst_count": 1})
    assert partial["available"] is True and partial["partial"] is True

    none = vs.calculate_neglect_score({})
    assert none["available"] is False and none["score"] is None


def test_neglect_bonus_requires_cheapness():
    """소외 가산점은 '싸고 소외된' 종목에만 붙는다."""
    import value_screener as vs

    high = {"available": True, "score": 90}
    assert vs.neglect_bonus(high, valuation_score=80) > 0    # 싸고 소외 → 가산
    assert vs.neglect_bonus(high, valuation_score=40) == 0   # 비싸면 가산 없음
    assert vs.neglect_bonus({"available": True, "score": 10}, 80) < 0  # 인기주는 감점
    assert vs.neglect_bonus({"available": False}, 80) == 0
    # 가산점 범위는 ±5로 제한
    assert abs(vs.neglect_bonus({"available": True, "score": 100}, 80)) <= 5


# ====================================================================
# 4. 가치 훼손 모니터
# ====================================================================
def _snapshot(price=10000, per=10.0, pbr=1.0, **kw):
    snap = {"name": "테스트", "price": price, "per": per, "pbr": pbr,
            "per_basis": "trailing",
            "eps": price / per, "bps": price / pbr,
            "roe_pct": 15.0, "operating_margin_pct": 12.0,
            "debt_to_equity": 50.0, "revenue_growth_pct": 8.0,
            "trap_flags": [], "asof": "2026-07-01"}
    snap.update(kw)
    return snap


def test_price_drop_alone_is_opportunity_not_risk():
    """브리핑의 핵심: 가치가 유지된 주가 하락은 위험이 아니라 기회."""
    import value_erosion_monitor as vem

    baseline = _snapshot(price=10000, per=10.0, pbr=1.0)
    # 주가만 -20%. 이익·순자산 그대로이므로 배수도 함께 내려감
    current = _snapshot(price=8000, per=8.0, pbr=0.8)

    signals = vem.detect_erosion_signals(current, baseline)
    assert signals == []

    status, _ = vem.classify(-20.0, 0)
    assert status == "OPPORTUNITY"


def test_earnings_erosion_flagged_even_when_price_holds():
    """주가가 버텨도 이익이 무너지면 경보 (훼손 은폐)."""
    import value_erosion_monitor as vem

    baseline = _snapshot(price=10000, per=10.0)      # EPS 1000
    current = _snapshot(price=10500, per=15.0)       # EPS 700 (-30%)

    signals = vem.detect_erosion_signals(current, baseline)
    types = [s["type"] for s in signals]
    assert "이익훼손" in types
    score = sum(s["severity"] for s in signals)

    status, _ = vem.classify(5.0, score)
    assert status == "EROSION_MASKED"


def test_erosion_plus_price_drop_is_alert():
    import value_erosion_monitor as vem

    baseline = _snapshot(price=10000, per=10.0, operating_margin_pct=12.0)
    current = _snapshot(price=7000, per=14.0, operating_margin_pct=6.0)

    signals = vem.detect_erosion_signals(current, baseline)
    score = sum(s["severity"] for s in signals)
    assert score >= vem.EROSION_THRESHOLD

    status, interpretation = vem.classify(-30.0, score)
    assert status == "EROSION"
    assert "매도 검토" in interpretation


def test_per_basis_switch_suppresses_false_earnings_signal():
    """후행 PER ↔ 선행 PER 교체로 튄 EPS는 훼손으로 보지 않는다."""
    import value_erosion_monitor as vem

    baseline = _snapshot(price=10000, per=10.0, per_basis="trailing")
    current = _snapshot(price=10000, per=20.0, per_basis="forward")

    types = [s["type"] for s in vem.detect_erosion_signals(current, baseline)]
    assert "이익훼손" not in types

    # 기준이 같으면 정상적으로 잡힌다
    current["per_basis"] = "trailing"
    types = [s["type"] for s in vem.detect_erosion_signals(current, baseline)]
    assert "이익훼손" in types


def test_other_erosion_signals():
    """수익성·재무·성장·함정·컨센서스 신호."""
    import value_erosion_monitor as vem

    baseline = _snapshot(roe_pct=20.0, debt_to_equity=50.0,
                         revenue_growth_pct=10.0, trap_flags=[])
    current = _snapshot(roe_pct=12.0, debt_to_equity=95.0,
                        revenue_growth_pct=-4.0, trap_flags=["사양산업"])

    types = {s["type"] for s in vem.detect_erosion_signals(current, baseline,
                                                           target_change_pct=-18.0)}
    assert {"ROE하락", "부채증가", "성장역전", "함정신규", "컨센서스컷"} <= types


def test_new_ticker_records_baseline_only():
    """최초 관측 종목은 판정하지 않고 기준선만 세운다."""
    import value_erosion_monitor as vem

    stocks = [{"ticker": "005930", "name": "삼성전자", "current_price": 80000,
               "metrics": {"per": 12.0, "pbr": 1.2, "per_basis": "trailing"}}]
    results, baselines = vem.analyze(stocks, {"baselines": {}}, {})

    assert results[0]["status"] == "NEW"
    assert "005930" in baselines
    assert baselines["005930"]["eps"] == 80000 / 12.0


def test_baseline_not_rebased_before_interval():
    """기준선은 REBASE_DAYS 전에는 교체되지 않는다.

    매 실행마다 갱신하면 서서히 진행되는 훼손이 매번 '변화 없음'으로 묻힌다.
    """
    import value_erosion_monitor as vem

    old = _snapshot(price=10000, per=10.0, asof="2026-07-30")
    stocks = [{"ticker": "005930", "name": "삼성전자", "current_price": 9000,
               "metrics": {"per": 9.0, "pbr": 0.9, "per_basis": "trailing"}}]

    _, baselines = vem.analyze(stocks, {"baselines": {"005930": old}}, {},
                               rebase_days=30, today=datetime(2026, 8, 2, tzinfo=vem.KST))
    assert baselines["005930"]["price"] == 10000        # 3일 경과 → 유지

    _, baselines = vem.analyze(stocks, {"baselines": {"005930": old}}, {},
                               rebase_days=30, today=datetime(2026, 9, 15, tzinfo=vem.KST))
    assert baselines["005930"]["price"] == 9000         # 47일 경과 → 교체


def test_message_empty_without_alerts():
    import value_erosion_monitor as vem

    assert vem.build_message([]) == ""


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
