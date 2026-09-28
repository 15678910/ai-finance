"""machine_economy_monitor.py — 파서·이력·판정 테스트 (네트워크는 모킹).

실행: python -m pytest tests/test_machine_economy.py -v
"""

import os
import sys
import json

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

M = pytest.importorskip("machine_economy_monitor")


# ====================================================================
# 파서
# ====================================================================
def test_parse_coingecko_markets():
    payload = [
        {"id": "tether", "symbol": "usdt", "market_cap": 180e9, "total_volume": 90e9, "price_change_percentage_24h": 0.01},
        {"id": "usd-coin", "symbol": "usdc", "market_cap": 75e9, "total_volume": 12e9, "price_change_percentage_24h": None},
        {"id": "junk", "market_cap": 0},
    ]
    r = M.parse_coingecko_markets(payload)
    assert r["source"] == "coingecko"
    assert r["total_cap_b"] == pytest.approx(255.0)
    assert r["total_vol24h_b"] == pytest.approx(102.0)
    assert [c["name"] for c in r["coins"]] == ["USDT", "USDC"]
    assert M.parse_coingecko_markets({"error": "rate limited"}) == {}
    assert M.parse_coingecko_markets([]) == {}


def test_parse_llama_stables():
    payload = {"peggedAssets": [
        {"symbol": "USDT", "gecko_id": "tether", "circulating": {"peggedUSD": 180e9}},
        {"symbol": "USDC", "gecko_id": "usd-coin", "circulating": {"peggedUSD": 75e9}},
        {"symbol": "X", "circulating": {"peggedUSD": 0}},
    ]}
    r = M.parse_llama_stables(payload)
    assert r["source"] == "defillama" and r["total_cap_b"] == pytest.approx(255.0)
    assert r["total_vol24h_b"] is None                       # 거래량 없음 — 정직하게 None
    assert M.parse_llama_stables({}) == {} and M.parse_llama_stables(None) == {}


def test_parse_vast_offers():
    payload = {"offers": [{"dph_total": 2.1}, {"dph_total": 1.9}, {"dph_total": 2.5},
                          {"dph_total": 9999}, {"dph_base": 2.0}, {"dph_total": "bad"}]}
    r = M.parse_vast_offers(payload)
    assert r["n_offers"] == 4 and r["min_usd_hr"] == 1.9
    assert r["median_usd_hr"] == pytest.approx(2.05)         # 9999 는 이상치로 제외
    assert M.parse_vast_offers({"offers": []}) == {}
    assert M.parse_vast_offers("<html>blocked</html>") == {}


def test_parse_x402_stats_accepts_common_shapes():
    assert M.parse_x402_stats({"total_transactions": "12345", "total_volume_usd": 6.7e6})["total_tx"] == 12345
    assert M.parse_x402_stats({"data": {"stats": {"txCount": 10, "volume": 5}}}) == {"total_tx": 10, "total_volume_usd": 5}
    assert M.parse_x402_stats({"unrelated": 1}) == {}
    assert M.parse_x402_stats("nope") == {}


# ====================================================================
# 이력·파생
# ====================================================================
def _series(days, start=100.0, step=1.0, key="stable_cap_b", first="2026-01-01"):
    from datetime import datetime, timedelta
    d0 = datetime.strptime(first, "%Y-%m-%d")
    return [{"date": (d0 + timedelta(days=i)).strftime("%Y-%m-%d"), key: start + step * i} for i in range(days)]


def test_merge_history_replaces_same_date_and_caps():
    s = _series(3)
    s2 = M.merge_history(s, {"date": "2026-01-03", "stable_cap_b": 999.0})
    assert len(s2) == 3 and s2[-1]["stable_cap_b"] == 999.0
    big = _series(M.MAX_POINTS + 50)
    assert len(M.merge_history(big, {"date": "2030-01-01", "stable_cap_b": 1})) == M.MAX_POINTS


def test_derive_trend_requires_real_lookback():
    """이력이 3일뿐이면 '30일 변화율'을 어제 값으로 계산하면 안 된다."""
    t = M.derive_trend(_series(3), "stable_cap_b", "2026-01-03")
    assert "chg_30d_pct" not in t and "annualized_growth_pct" not in t
    assert t["current"] == 102.0


def test_derive_trend_with_enough_history():
    s = _series(120, start=100.0, step=0.5)               # 120일간 +60
    today = s[-1]["date"]
    t = M.derive_trend(s, "stable_cap_b", today)
    assert t["chg_30d_pct"] == pytest.approx((159.5 / 144.5 - 1) * 100, rel=1e-3)
    assert t["chg_90d_pct"] is not None and "chg_365d_pct" not in t
    assert t["annualized_growth_pct"] > 100                   # 4개월에 +59.5% → 연환산 세 자릿수


def test_derive_trend_missing_key():
    assert M.derive_trend(_series(5), "h100_usd_hr", "2026-01-05") == {}
    assert M.derive_trend([], "stable_cap_b", "2026-01-05") == {}


# ====================================================================
# 판정
# ====================================================================
def test_claims_status_labels():
    stable = {"total_cap_b": 300.0}
    compute = {"h100": {"median_usd_hr": 2.0, "min_usd_hr": 1.8, "n_offers": 50}}
    trend = {"stable_cap_b": {"annualized_growth_pct": 65.0, "days_of_history": 200},
             "h100_usd_hr": {"chg_90d_pct": -15.0}}
    c = M.build_claims(stable, compute, {}, trend)
    assert [x["status"] for x in c] == ["proxy", "measured", "unavailable"]
    assert "근접" in c[0]["verdict"] and "하락" in c[1]["verdict"] and "미확보" in c[2]["verdict"]


def test_claims_when_everything_fails():
    c = M.build_claims({}, {}, {}, {})
    assert all(x["status"] == "unavailable" for x in c)


def test_claims_x402_measured_when_available():
    c = M.build_claims({"total_cap_b": 1}, {}, {"total_tx": 500, "source": "x402scan.com"}, {})
    assert c[2]["status"] == "measured" and "500" in c[2]["detail"]


# ====================================================================
# 끝까지 — 실패 시 직전 값 유지 + 파일 계약
# ====================================================================
def test_main_keeps_last_values_when_sources_fail(monkeypatch, tmp_path):
    monkeypatch.setattr(M, "OUTPUT_FILE", str(tmp_path / "machine_economy.json"))
    monkeypatch.setattr(M, "load_state", lambda name, default=None: {"series": [
        {"date": "2026-09-27", "stable_cap_b": 310.0, "stable_vol_b": 100.0, "h100_usd_hr": 2.2}]})
    saved = {}
    monkeypatch.setattr(M, "save_state", lambda name, data: saved.update(data) or True)
    monkeypatch.setattr(M, "fetch_stablecoins", lambda: {})
    monkeypatch.setattr(M, "fetch_compute", lambda: {})
    monkeypatch.setattr(M, "fetch_x402", lambda: {})
    assert M.main() == 0
    out = json.load(open(tmp_path / "machine_economy.json", encoding="utf-8"))
    assert out["facts"]["stale"] == ["stablecoins", "compute"]
    assert out["facts"]["stablecoins"]["total_cap_b"] == 310.0
    assert out["facts"]["stablecoins"]["source"].startswith("stale(")
    assert len(saved["series"]) == 1                          # 전부 stale 이면 새 점을 쌓지 않는다
    assert set(out) >= {"generated_at", "facts", "trend", "claims", "thesis", "note", "history_tail"}


def test_main_appends_point_on_success(monkeypatch, tmp_path):
    monkeypatch.setattr(M, "OUTPUT_FILE", str(tmp_path / "me.json"))
    monkeypatch.setattr(M, "load_state", lambda name, default=None: {"series": []})
    saved = {}
    monkeypatch.setattr(M, "save_state", lambda name, data: saved.update(data) or True)
    monkeypatch.setattr(M, "fetch_stablecoins", lambda: {"total_cap_b": 305.0, "total_vol24h_b": 98.0, "coins": [], "source": "coingecko"})
    monkeypatch.setattr(M, "fetch_compute", lambda: {"h100": {"gpu": "H100 SXM", "median_usd_hr": 2.05, "min_usd_hr": 1.9, "n_offers": 40}, "source": "vast.ai"})
    monkeypatch.setattr(M, "fetch_x402", lambda: {})
    assert M.main() == 0
    assert len(saved["series"]) == 1 and saved["series"][0]["h100_usd_hr"] == 2.05
    out = json.load(open(tmp_path / "me.json", encoding="utf-8"))
    assert out["claims"][1]["status"] == "measured" and out["claims"][2]["status"] == "unavailable"
    assert out["thesis"]["source"].startswith("BlackRock")
