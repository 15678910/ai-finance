"""DCF 평가기 테스트 — 투자은행 실무 관행 반영분.

대상: dcf_valuator.py
  · 연중 할인 (mid-year convention)
  · 터미널 가치 이원화 (영구성장법 + EV/EBITDA 출구배수법)
  · 희석주식수 선택
  · 성장 경로 헬퍼 (리팩터 무손실)

실행: python -m pytest tests/test_dcf.py -v
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

D = pytest.importorskip("dcf_valuator")

WACC = 0.09


def _fcf():
    return D.project_fcf([100.0], 0.05)


# ====================================================================
# 성장 경로 · 리팩터 무손실
# ====================================================================
def test_growth_path_converges_to_terminal():
    path = D.growth_path(0.20, 5)
    assert len(path) == 5
    assert path[-1] == pytest.approx(D.TERMINAL_GROWTH)
    assert all(path[i] >= path[i + 1] for i in range(4))    # 단조 감소


def test_project_fcf_matches_previous_inline_formula():
    """growth_path 로 뽑아낸 뒤에도 옛 인라인 공식과 수치가 같아야 한다."""
    def old(hist, g, years=5):
        cur, out = hist[-1], []
        for y in range(1, years + 1):
            decay = (g - D.TERMINAL_GROWTH) * (1 - y / years)
            cur *= 1 + D.TERMINAL_GROWTH + decay
            out.append(cur)
        return out

    hist, g = [80.0, 90.0, 100.0], 0.12
    assert D.project_fcf(hist, g) == pytest.approx(old(hist, g))


# ====================================================================
# 연중 할인
# ====================================================================
def test_mid_year_raises_value_by_half_year_of_discount():
    """연중 할인은 모든 흐름을 반년 앞당기므로 EV가 (1+WACC)^0.5 만큼 커진다."""
    fcf = _fcf()
    end = D.calculate_dcf(fcf, WACC, mid_year=False)["enterprise_value"]
    mid = D.calculate_dcf(fcf, WACC, mid_year=True)["enterprise_value"]
    assert mid / end == pytest.approx((1 + WACC) ** 0.5, rel=1e-9)


def test_mid_year_is_default():
    fcf = _fcf()
    default = D.calculate_dcf(fcf, WACC)["enterprise_value"]
    explicit = D.calculate_dcf(fcf, WACC, mid_year=D.MID_YEAR_CONVENTION)["enterprise_value"]
    assert default == explicit


def test_terminal_share_reported():
    r = D.calculate_dcf(_fcf(), WACC)
    assert r["terminal_share_pct"] is not None
    assert 50 < r["terminal_share_pct"] < 95       # 5년 추정이면 터미널이 EV 대부분


# ====================================================================
# 터미널 가치 이원화
# ====================================================================
def test_exit_multiple_discounted_at_year_end_not_mid_year():
    """출구배수 터미널은 연말 매각 가정이라 N년 정수 할인. 연중으로 당기면 과대평가."""
    fcf = _fcf()
    n = len(fcf)
    r = D.calculate_dcf(fcf, WACC, mid_year=True, terminal_ebitda=200.0, exit_multiple=8.0)
    assert r["terminal_value_exit"] == pytest.approx(1600.0)
    assert r["pv_of_terminal_exit"] == pytest.approx(1600.0 / (1 + WACC) ** n)
    assert r["enterprise_value_exit"] == pytest.approx(r["pv_of_fcf"] + r["pv_of_terminal_exit"])


def test_exit_multiple_absent_leaves_gordon_only():
    r = D.calculate_dcf(_fcf(), WACC)
    assert r["enterprise_value"] > 0
    assert r["enterprise_value_exit"] is None
    assert r["pv_of_terminal_exit"] is None


def test_exit_multiple_ignored_when_ebitda_nonpositive():
    r = D.calculate_dcf(_fcf(), WACC, terminal_ebitda=0.0, exit_multiple=8.0)
    assert r["enterprise_value_exit"] is None
    r = D.calculate_dcf(_fcf(), WACC, terminal_ebitda=-50.0, exit_multiple=8.0)
    assert r["enterprise_value_exit"] is None


def test_select_exit_multiple_uses_market_within_range():
    lo, hi = D.EXIT_MULTIPLE_RANGE
    assert D.select_exit_multiple({"ev_to_ebitda": 11.3}) == (11.3, "market")
    assert D.select_exit_multiple({"ev_to_ebitda": hi + 20}) == (D.DEFAULT_EXIT_MULTIPLE, "default")
    assert D.select_exit_multiple({"ev_to_ebitda": lo - 1}) == (D.DEFAULT_EXIT_MULTIPLE, "default")
    assert D.select_exit_multiple({}) == (D.DEFAULT_EXIT_MULTIPLE, "default")
    assert D.select_exit_multiple({"ev_to_ebitda": "n/a"}) == (D.DEFAULT_EXIT_MULTIPLE, "default")


# ====================================================================
# 희석주식수
# ====================================================================
def test_select_share_count_prefers_implied_diluted():
    shares, basis = D.select_share_count({"shares_outstanding": 100, "shares_implied": 104})
    assert shares == 104 and basis == "implied_diluted"


def test_select_share_count_rejects_implied_outlier():
    """implied 가 basic 의 2배 밖이면 단위 오류 — 기본주식수로 복귀."""
    shares, basis = D.select_share_count({"shares_outstanding": 100, "shares_implied": 900})
    assert shares == 100 and basis.startswith("basic")
    shares, basis = D.select_share_count({"shares_outstanding": 100, "shares_implied": 30})
    assert shares == 100 and basis.startswith("basic")


def test_select_share_count_fallbacks():
    assert D.select_share_count({"shares_outstanding": 100}) == (100.0, "basic")
    assert D.select_share_count({"shares_implied": 100}) == (100.0, "implied_diluted")
    assert D.select_share_count({}) == (0.0, "none")


def test_diluted_shares_lower_fair_price():
    """같은 Equity Value 를 더 많은 주식으로 나누면 적정주가는 내려가야 한다."""
    equity = 1_000_000.0
    basic, _ = D.select_share_count({"shares_outstanding": 100})
    diluted, _ = D.select_share_count({"shares_outstanding": 100, "shares_implied": 110})
    assert equity / diluted < equity / basic


# ====================================================================
# 방어
# ====================================================================
def test_invalid_inputs_return_empty():
    assert D.calculate_dcf([], WACC)["enterprise_value"] == 0
    assert D.calculate_dcf(_fcf(), D.TERMINAL_GROWTH)["enterprise_value"] == 0
    assert D.calculate_dcf(_fcf(), 0.01)["enterprise_value"] == 0
