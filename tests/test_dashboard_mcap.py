"""generate_dashboard_data — 시가총액 통화 환산 헬퍼.

2026-09-20 실측: build_treemap_sectors() 가 원→달러 환산을 마친 뒤, 10.14 '가격 최신화'
단계가 info["marketCap"]/1e12 (원화 그대로) 로 다시 덮어써서 삼성전자가 1713.87조$
로 찍혔다(실제 1.24조$). 환산 규칙을 헬퍼 하나로 모아 두 곳이 같은 함수를 쓴다.

실행: python -m pytest tests/test_dashboard_mcap.py -v
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

G = pytest.importorskip("generate_dashboard_data")
f = G._mcap_trillion_usd


def test_krw_is_converted_with_fx():
    # 삼성전자 실측: 1,713.87조원 / 1385 ≈ 1.237조$
    assert f(1713.87e12, is_krw=True, fx=1385.0) == pytest.approx(1.237, abs=0.001)


def test_usd_and_crypto_pass_through():
    assert f(5.5e12, is_krw=False, fx=1385.0) == pytest.approx(5.5)
    assert f(5.5e12, is_krw=False, fx=None) == pytest.approx(5.5)


def test_krw_without_fx_returns_none_not_raw():
    """환율이 없으면 원화 값을 '조$'로 내보내면 안 된다 — 덮어쓰기 자체를 건너뛰도록 None."""
    assert f(1713.87e12, is_krw=True, fx=None) is None
    assert f(1713.87e12, is_krw=True, fx=0) is None


def test_invalid_inputs():
    assert f(None, is_krw=True, fx=1385.0) is None
    assert f(0, is_krw=False, fx=None) is None
    assert f("abc", is_krw=False, fx=None) is None


def test_rounding_keeps_small_caps_visible():
    # 소형 코인·종목이 0.0 으로 사라지지 않도록 소수 3자리
    assert f(4.2e9, is_krw=False, fx=None) == 0.004
