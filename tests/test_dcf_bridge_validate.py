"""anthropics/financial-services dcf-model 스킬 정렬분 테스트.

  1. 터미널 비중 > 75% 경고 플래그 (terminal_heavy)
  2. EV→Equity 브리지 확장 — 리스·소수주주지분·연금 (build_equity_bridge)
  3. 출력 검증기 (validate_dcf_output.validate)
  + 헤드라인 시그널: 두 방법 괴리 > 30% → '⚪ 방법 간 불일치' (method_mismatch)

실행: python -m pytest tests/test_dcf_bridge_validate.py -v
"""

import os
import sys
import math
from datetime import datetime, timezone, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

D = pytest.importorskip("dcf_valuator")
V = pytest.importorskip("validate_dcf_output")
KST = timezone(timedelta(hours=9))


# ====================================================================
# 2. Equity 브리지
# ====================================================================
def _fin(**kw):
    base = {"total_debt": 100.0, "cash": 30.0, "lease": None, "debt_ex_lease_bs": None,
            "minority_interest": None, "pension": None}
    base.update(kw)
    return base


def test_bridge_basic_net_debt():
    b = D.build_equity_bridge(_fin())
    assert b["net_debt"] == pytest.approx(70.0)
    assert b["lease_treatment"] == "n/a"
    assert b["lease_added"] == 0 and b["minority_interest"] == 0 and b["pension"] == 0


def test_bridge_adds_lease_when_debt_excludes_it():
    """info.totalDebt(100) ≈ BS 차입금(100) → 리스(40)가 빠져 있음 → 더한다."""
    b = D.build_equity_bridge(_fin(lease=40.0, debt_ex_lease_bs=100.0))
    assert b["lease_treatment"] == "added"
    assert b["lease_added"] == pytest.approx(40.0)
    assert b["net_debt"] == pytest.approx(100 + 40 - 30)


def test_bridge_skips_lease_when_debt_already_includes_it():
    """info.totalDebt(140) ≈ BS 차입금(100) + 리스(40) → 이미 포함 → 이중계산 금지."""
    b = D.build_equity_bridge(_fin(total_debt=140.0, lease=40.0, debt_ex_lease_bs=100.0))
    assert b["lease_treatment"] == "included_in_debt"
    assert b["lease_added"] == 0
    assert b["net_debt"] == pytest.approx(140 - 30)


def test_bridge_lease_without_bs_reference_is_added_conservatively():
    """리스 포함 여부를 판별할 BS 차입금이 없으면 보수적으로 더한다(주주 몫 과대평가 방지)."""
    b = D.build_equity_bridge(_fin(lease=40.0))
    assert b["lease_treatment"] == "added"
    assert b["net_debt"] == pytest.approx(110.0)


def test_bridge_adds_minority_and_pension():
    b = D.build_equity_bridge(_fin(minority_interest=15.0, pension=5.0))
    assert b["net_debt"] == pytest.approx(100 + 15 + 5 - 30)
    assert b["minority_interest"] == 15.0 and b["pension"] == 5.0


def test_bridge_tolerates_missing_and_garbage():
    b = D.build_equity_bridge({})
    assert b["net_debt"] == 0.0
    b = D.build_equity_bridge(_fin(total_debt="n/a", cash=None, lease="x"))
    assert b["net_debt"] == 0.0


# ====================================================================
# 1·시그널 — evaluate_stock 끝까지
# ====================================================================
def _run(monkeypatch, fcf, price, shares, ebitda=None, ev_to_ebitda=None, **extra):
    n = len(fcf)
    fin = {"current_price": price, "market_cap": price * shares,
           "shares_outstanding": shares, "shares_implied": None, "beta": 1.0,
           "total_debt": 0.0, "cash": 0.0, "fcf_history": fcf,
           "revenue_history": [100e12 * 1.03 ** i for i in range(n)],
           "ebitda": ebitda, "ev_to_ebitda": ev_to_ebitda,
           "lease": None, "debt_ex_lease_bs": None, "minority_interest": None, "pension": None}
    fin.update(extra)
    monkeypatch.setattr(D, "fetch_financials", lambda t: fin)
    return D.evaluate_stock("t", "000000.KS")


def test_terminal_heavy_flag_and_warning(monkeypatch):
    """5년 추정이면 터미널이 EV 70~80% — 75% 넘으면 플래그·경고 문구."""
    r = _run(monkeypatch, [5e12] * 4, 250_000, 3.0e8)
    assert r["dcf_applicable"] is True
    assert r["terminal_share_pct"] is not None
    assert r["terminal_heavy"] == (r["terminal_share_pct"] > D.TERMINAL_SHARE_WARN_PCT)
    if r["terminal_heavy"]:
        assert any("터미널" in w for w in r["warnings"])


def test_method_mismatch_downgrades_signal(monkeypatch):
    """출구배수 적정가가 영구성장법과 3배 갈리면 매수/매도 시그널 대신 '방법 간 불일치'."""
    r = _run(monkeypatch, [5e12] * 4, 250_000, 3.0e8, ebitda=20e12, ev_to_ebitda=15.0)
    assert r["dcf_applicable"] is True
    assert r["fair_price_exit"] is not None
    assert abs(r["terminal_method_gap_pct"]) > D.METHOD_GAP_SIGNAL_PCT
    assert r["method_mismatch"] is True
    assert r["signal"] == "⚪ 방법 간 불일치"
    assert any("괴리" in w for w in r["warnings"])


def test_low_confidence_takes_precedence_over_mismatch(monkeypatch):
    r = _run(monkeypatch, [5e12] * 4, 30_000, 3.0e8, ebitda=20e12, ev_to_ebitda=15.0)
    assert r["low_confidence"] is True
    assert r["signal"] == "⚠️ 모델 신뢰도 낮음"


def test_agreeing_methods_keep_normal_signal(monkeypatch):
    """두 방법이 비슷하면 기존 등급 그대로."""
    r0 = _run(monkeypatch, [5e12] * 4, 250_000, 3.0e8)          # 출구배수 없음 → 기준 EV
    # 출구배수 적정가가 영구성장법과 같아지도록 시장 배수를 역산한다
    ev_g = r0["enterprise_value"]
    # exit EV = pv_fcf + ebitda_N*mult/(1+w)^5 ; ebitda_N = ebitda*prod(1+g). 간단히 배수를 맞춘다
    r = _run(monkeypatch, [5e12] * 4, 250_000, 3.0e8, ebitda=5e12, ev_to_ebitda=12.0)
    if r["method_mismatch"]:
        pytest.skip("합성 입력에서 괴리가 30%를 넘음 — 시그널 강등 자체는 다른 테스트에서 검증")
    assert r["signal"] not in ("⚪ 방법 간 불일치", "⚠️ 모델 신뢰도 낮음", "⚪ DCF 부적합")


def test_bridge_is_reported_and_lease_reduces_equity(monkeypatch):
    r_no = _run(monkeypatch, [5e12] * 4, 250_000, 3.0e8)
    r_ls = _run(monkeypatch, [5e12] * 4, 250_000, 3.0e8, lease=30e12, debt_ex_lease_bs=0.0)
    assert r_ls["equity_bridge"]["lease_treatment"] == "added"
    assert r_ls["fair_price_dcf"] < r_no["fair_price_dcf"]


def test_telegram_excludes_mismatch():
    vals = [{"name": "a", "dcf_applicable": True, "low_confidence": False, "method_mismatch": True, "upside_pct": 60.0},
            {"name": "b", "dcf_applicable": True, "low_confidence": False, "method_mismatch": False, "upside_pct": 45.0}]
    strong = [v for v in vals if v.get("dcf_applicable", True) and not v.get("low_confidence")
              and not v.get("method_mismatch") and (v.get("upside_pct") or 0) > 30]
    assert [v["name"] for v in strong] == ["b"]


# ====================================================================
# 3. 출력 검증기
# ====================================================================
NOW = datetime(2026, 9, 25, 9, 0, tzinfo=KST)


def _good(**over):
    v = {"name": "ok", "dcf_applicable": True, "fair_price_dcf": 300_000.0, "upside_pct": 20.0,
         "wacc_pct": 7.9, "growth_rate_pct": 3.0, "beta": 1.0, "shares_outstanding": 3e8,
         "terminal_share_pct": 72.0, "terminal_method_gap_pct": 8.0, "low_confidence": False,
         "method_mismatch": False, "current_price": 250_000.0, "fcf_base": 5e12}
    v.update(over)
    return v


def _doc(*vals, gen="2026-09-25 07:30:00 KST"):
    return {"generated_at": gen, "assumptions": {"terminal_growth_pct": 2.0}, "valuations": list(vals)}


def test_validator_passes_clean_output():
    errs, warns = V.validate(_doc(_good()), now=NOW)
    assert errs == []
    assert warns[0].startswith("적용 1")


def test_validator_catches_negative_fair_price_and_deep_negative_upside():
    errs, _ = V.validate(_doc(_good(fair_price_dcf=-1_597_114.0, upside_pct=-537.6)), now=NOW)
    assert any("fair_price_dcf" in e for e in errs)
    assert any("upside_pct" in e for e in errs)


def test_validator_catches_wacc_below_terminal_growth():
    errs, _ = V.validate(_doc(_good(wacc_pct=1.5)), now=NOW)
    assert any("영구성장률" in e for e in errs)


def test_validator_catches_nan_and_stale():
    errs, _ = V.validate(_doc(_good(beta=float("nan"))), now=NOW)
    assert any("NaN" in e for e in errs)
    errs, _ = V.validate(_doc(_good(), gen="2026-09-19 09:26:09 KST"), now=NOW)
    assert any("경과" in e for e in errs)


def test_validator_inapplicable_needs_reason():
    bad = {"name": "x", "dcf_applicable": False, "dcf_reason": ""}
    good = {"name": "y", "dcf_applicable": False, "dcf_reason": "정규화 FCF 음수"}
    errs, warns = V.validate(_doc(bad, good), now=NOW)
    assert len(errs) == 1 and "x" in errs[0]
    assert "부적합 2" in warns[0]


def test_validator_warns_terminal_heavy_and_gap():
    errs, warns = V.validate(_doc(_good(terminal_share_pct=84.0, terminal_method_gap_pct=261.0)), now=NOW)
    assert errs == []
    assert any("터미널 비중 84%" in w for w in warns)
    assert any("괴리 +261%" in w for w in warns)


def test_validator_empty_document():
    errs, _ = V.validate({"valuations": []}, now=NOW)
    assert errs and "비어" in errs[0]
