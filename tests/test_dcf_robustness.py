"""DCF 평가기 견고성 테스트 — 비현실적 적정가(-250%) 재발 방지.

대상: dcf_valuator.py
  · FCF 정규화 (중앙값) 와 DCF 부적합 판정
  · 베타 Blume 보정·클램프
  · project_fcf 의 base_fcf 인자

실행: python -m pytest tests/test_dcf_robustness.py -v
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

D = pytest.importorskip("dcf_valuator")

# 2026-09-19 실측 FCF 이력 (조원, 오래된→최근)
HYUNDAI = [4.9e12, -11.4e12, -16.0e12, -17.1e12]      # 금융자회사 연결 → 최근 음수
LGES = [-6.9e12, -5.6e12, -7.4e12, -6.6e12]            # 설비투자 사이클 → 전부 음수
SAMSUNG = [9.1e12, -16.4e12, 19.2e12, 33.2e12]         # 한 해 변동 큼
KIA = [7.2e12, 8.2e12, 7.9e12, 4.1e12]


# ====================================================================
# FCF 정규화
# ====================================================================
def test_normalize_fcf_uses_median_not_latest():
    base, basis = D.normalize_fcf(SAMSUNG)
    assert basis == "4년 중앙값"
    assert base == pytest.approx((9.1e12 + 19.2e12) / 2)     # 극단값(-16.4, 33.2) 제외
    assert base != SAMSUNG[-1]


def test_normalize_fcf_negative_cycle_is_negative():
    """음수 사이클 회사는 정규화해도 음수 → 부적합 판정으로 이어져야 한다."""
    assert D.normalize_fcf(LGES)[0] < 0
    assert D.normalize_fcf(HYUNDAI)[0] < 0


def test_normalize_fcf_stable_company_close_to_latest():
    base, _ = D.normalize_fcf(KIA)
    assert 7.0e12 < base < 8.5e12


def test_normalize_fcf_edge_cases():
    assert D.normalize_fcf([]) == (0.0, "none")
    assert D.normalize_fcf(None) == (0.0, "none")
    base, basis = D.normalize_fcf([5.0])
    assert base == 5.0 and basis == "1년 중앙값"
    # N년 초과분은 오래된 쪽을 버린다
    base, basis = D.normalize_fcf([100.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
    assert base == 1.0 and basis == f"{D.FCF_NORMALIZE_YEARS}년 중앙값"


# ====================================================================
# 음수 기준 FCF 가 더 이상 음수 적정가를 만들지 않는다
# ====================================================================
def test_negative_base_never_projected():
    """옛 코드는 음수 FCF 를 성장시켜 더 큰 음수로 만들었다. 새 경로는 기준점을 갈아끼운다."""
    proj_old_style = D.project_fcf(HYUNDAI, 0.10)              # 하위 호환: 최근값에서 출발
    assert all(v < 0 for v in proj_old_style)                  # 그래서 이 경로는 쓰지 않는다
    proj_new = D.project_fcf(HYUNDAI, 0.10, base_fcf=8.0e12)   # 정규화 기준점 명시
    assert all(v > 0 for v in proj_new)
    assert proj_new[0] == pytest.approx(8.0e12 * (1 + D.growth_path(0.10)[0]))


def test_project_fcf_base_fcf_overrides_history():
    a = D.project_fcf([1.0, 2.0, 99.0], 0.05, base_fcf=10.0)
    b = D.project_fcf([10.0], 0.05)
    assert a == pytest.approx(b)


# ====================================================================
# 베타 보정
# ====================================================================
def test_adjust_beta_pulls_toward_one():
    assert D.adjust_beta(1.54) == pytest.approx(0.67 * 1.54 + 0.33)   # 삼성전자 실측
    assert 1.0 < D.adjust_beta(1.54) < 1.54
    assert D.adjust_beta(1.0) == pytest.approx(1.0)


def test_adjust_beta_clamps_extremes():
    lo, hi = D.BETA_RANGE
    assert D.adjust_beta(0.21) == lo          # 삼성바이오로직스 실측 → 하한
    assert D.adjust_beta(3.5) == hi
    assert D.adjust_beta(0.0) == pytest.approx(1.0)
    assert D.adjust_beta(None) == pytest.approx(1.0)
    assert D.adjust_beta("n/a") == pytest.approx(1.0)


def test_adjusted_beta_narrows_wacc_spread():
    """실측 극단 두 종목의 WACC 격차가 보정 후 줄어야 한다."""
    raw = abs(D.calculate_wacc(1.54) - D.calculate_wacc(0.21))
    adj = abs(D.calculate_wacc(D.adjust_beta(1.54)) - D.calculate_wacc(D.adjust_beta(0.21)))
    assert adj < raw * 0.75


# ====================================================================
# 끝까지 돌려서 확인 — 실측 입력으로 evaluate_stock 경로 전체
# ====================================================================
def _run(monkeypatch, name, fcf_history, price, shares, net_debt=0.0, beta=1.0, cagr=0.05):
    n = len(fcf_history)
    fin = {"current_price": price, "market_cap": price * shares,
           "shares_outstanding": shares, "shares_implied": None, "beta": beta,
           "total_debt": max(net_debt, 0.0), "cash": max(-net_debt, 0.0),
           "fcf_history": fcf_history,
           "revenue_history": [100e12 * (1 + cagr) ** i for i in range(n)],
           "ebitda": None, "ev_to_ebitda": None}
    monkeypatch.setattr(D, "fetch_financials", lambda t: fin)
    return D.evaluate_stock(name, "000000.KS")


def test_negative_fcf_company_marked_inapplicable(monkeypatch):
    r = _run(monkeypatch, "현대차", HYUNDAI, 200_000, 2.02e8, net_debt=100e12, beta=1.73)
    assert r["dcf_applicable"] is False
    assert r["fair_price_dcf"] is None and r["upside_pct"] is None
    assert "음수" in r["dcf_reason"]
    assert r["signal"] == "⚪ DCF 부적합"
    assert r["fcf_base"] < 0


def test_no_negative_fair_price_ever(monkeypatch):
    """어떤 입력이든 음수 적정가·-100% 미만 괴리는 나오지 않아야 한다."""
    for hist in (HYUNDAI, LGES, [-1.0e12], [0.0, 0.0]):
        r = _run(monkeypatch, "x", hist, 10_000, 1e8)
        assert r["dcf_applicable"] is False
    r = _run(monkeypatch, "y", [5e12] * 4, 10_000, 1e8, net_debt=9.99e15)   # 순부채 > EV
    assert r["dcf_applicable"] is False and "순부채" in r["dcf_reason"]


def test_extreme_upside_flagged_not_signaled(monkeypatch):
    r = _run(monkeypatch, "기아", KIA, 30_000, 3.89e8, net_debt=-21e12, beta=1.01, cagr=0.097)
    assert r["dcf_applicable"] is True
    assert r["upside_pct"] > 100
    assert r["low_confidence"] is True
    assert "신뢰도" in r["signal"] and "매수" not in r["signal"]


def test_normal_case_keeps_buy_signal(monkeypatch):
    """FCF 5조 · 3억주 → 적정가 ≈ 30만원 수준. 현재가 25만원이면 온건한 저평가."""
    r = _run(monkeypatch, "z", [5e12] * 4, 250_000, 3.0e8, beta=1.0, cagr=0.03)
    assert r["dcf_applicable"] is True
    assert abs(r["upside_pct"]) < D.EXTREME_UPSIDE_PCT
    assert r["low_confidence"] is False
    assert r["signal"] not in ("⚪ DCF 부적합", "⚠️ 모델 신뢰도 낮음")
    assert r["fcf_basis"] == "4년 중앙값"


def test_short_fcf_history_is_inapplicable_card(monkeypatch):
    """1년치 FCF 는 None 이 아니라 '부적합' 카드로 남아야 한다 (카드 증발 방지)."""
    r = _run(monkeypatch, "w", [3e12], 10_000, 1e8)
    assert r is not None
    assert r["dcf_applicable"] is False
    assert "이력 부족" in r["dcf_reason"]


def test_telegram_excludes_inapplicable_and_low_confidence():
    """None upside 로 정렬·필터가 터지지 않고, 경고 종목은 강한매수에서 빠진다."""
    vals = [
        {"name": "a", "dcf_applicable": False, "upside_pct": None},
        {"name": "b", "dcf_applicable": True, "low_confidence": True, "upside_pct": 270.0},
        {"name": "c", "dcf_applicable": True, "low_confidence": False, "upside_pct": 45.0},
        {"name": "d", "dcf_applicable": True, "low_confidence": False, "upside_pct": 10.0},
    ]
    vals.sort(key=lambda x: (x.get("dcf_applicable", True), x.get("upside_pct") or -1e9), reverse=True)
    assert [v["name"] for v in vals] == ["b", "c", "d", "a"]         # 부적합은 맨 뒤
    strong = [v for v in vals if v.get("dcf_applicable", True)
              and not v.get("low_confidence") and (v.get("upside_pct") or 0) > 30]
    assert [v["name"] for v in strong] == ["c"]
