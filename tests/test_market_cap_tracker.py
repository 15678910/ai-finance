"""market_cap_tracker.py — 파서·이력·보정 로직 테스트 (네트워크는 모킹).

실행: python -m pytest tests/test_market_cap_tracker.py -v
"""

import os
import sys
import json
from datetime import datetime, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

M = pytest.importorskip("market_cap_tracker")


# ====================================================================
# 파서
# ====================================================================
def test_parse_naver_mv_page_shapes():
    p = {"stocks": [{"itemCode": "005930", "marketValue": "4,954,163"},
                    {"itemCode": "000660", "marketCap": 2500000},
                    {"itemCode": "bad", "marketValue": "1"},
                    {"itemCode": "123456", "marketValue": "0"}]}
    assert M.parse_naver_mv_page(p) == [("005930", 4954163.0), ("000660", 2500000.0)]
    assert M.parse_naver_mv_page([{"cd": "035420", "marketValue": "300,000"}]) == [("035420", 300000.0)]
    assert M.parse_naver_mv_page({"error": 1}) == [] and M.parse_naver_mv_page("html") == []


def test_sum_market_pages_dedupes_and_converts_to_trillion_won():
    pages = [[("005930", 4_000_000), ("000660", 2_000_000)], [("000660", 2_000_000), ("035420", 400_000)]]
    total, n = M.sum_market_pages(pages)
    assert n == 3 and total == pytest.approx(640.0)             # 6,400,000억 = 640조


def test_parse_coingecko_global():
    r = M.parse_coingecko_global({"data": {"total_market_cap": {"usd": 3.9e12}, "market_cap_change_percentage_24h_usd": -1.23,
                                           "market_cap_percentage": {"btc": 58.2}}})
    assert r == {"cap_t": 3.9, "chg_24h_pct": -1.23, "btc_dominance": 58.2}
    assert M.parse_coingecko_global({"data": {}}) == {} and M.parse_coingecko_global([]) == {}


def test_parse_sp500_csv_requires_plausible_count():
    rows = "Symbol,Security\n" + "\n".join(f"T{i},Co{i}" for i in range(500)) + "\nBRK.B,Berkshire\n"
    out = M.parse_sp500_csv(rows)
    assert len(out) == 501 and "BRK-B" in out
    assert M.parse_sp500_csv("Symbol\nAAPL\nMSFT\n") == []                 # 2개 → 목록 아님
    assert M.parse_sp500_csv("") == []


def test_parse_wiki_tickers():
    html = "<table>" + "".join(f"<tr><td>Company</td><td>{t}</td></tr>" for t in
                              [chr(65 + i // 26) + chr(65 + i % 26) + "X" for i in range(100)]) + "</table>"
    out = M.parse_wiki_tickers(html, (95, 110))
    assert len(out) == 100 and out[0] == "AAX"
    assert M.parse_wiki_tickers(html, (480, 530)) == []                    # 개수 범위 밖 → 거부
    assert M.parse_wiki_tickers("", (1, 10)) == []


# ====================================================================
# 이력
# ====================================================================
def _pts(n, start="2026-08-01T15:30", step_h=24, v0=100.0):
    t0 = datetime.fromisoformat(start)
    return [{"t": (t0 + timedelta(hours=step_h * i)).isoformat(timespec="minutes"), "v": v0 + i} for i in range(n)]


def test_append_point_replaces_same_timestamp():
    s = M.append_point([], "2026-09-29T10:00", 100)
    s = M.append_point(s, "2026-09-29T10:00", 101, est=False)
    s = M.append_point(s, "2026-09-29T11:00", 102)
    assert [p["v"] for p in s] == [101.0, 102.0] and "est" not in s[0]


def test_thin_series_keeps_recent_all_and_older_daily():
    now = "2026-09-29T12:00"
    old = _pts(48, start="2026-09-01T00:00", step_h=1)                    # 9/1~9/2 시간별 48점 → 2점
    recent = _pts(30, start="2026-09-28T00:00", step_h=1)                 # 최근 7일 안 → 전부
    out = M.thin_series(old + recent, now)
    assert len(out) == 2 + 30
    assert out[0]["t"].startswith("2026-09-01T23") and out[1]["t"].startswith("2026-09-02T23")
    assert len(M.thin_series(_pts(2000, step_h=1, start="2026-09-25T00:00"), "2026-12-31T00:00", max_points=50)) == 50


def test_change_1d_uses_point_at_least_20h_old():
    s = _pts(3, start="2026-09-27T15:30", step_h=24)                       # 9/27 100, 9/28 101, 9/29 102
    s.append({"t": "2026-09-29T16:00", "v": 103.0})                         # 같은 날 장중 점
    assert M.change_1d(s, "2026-09-29T16:00") == pytest.approx((103 / 101 - 1) * 100, abs=0.01)
    assert M.change_1d(_pts(1), "2026-08-01T16:00") is None
    assert M.change_1d([], "2026-08-01T16:00") is None


def test_backfill_and_scale():
    bf = M.backfill_from_index(60.0, [("2026-01-02", 5000.0), ("2026-01-03", 5100.0)], 6000.0)
    assert [p["v"] for p in bf] == [50.0, 51.0] and all(p["est"] for p in bf)
    assert M.backfill_from_index(None, [("2026-01-02", 1)], 1) == []
    assert M.scale_calibrated({"cap_t": 60.0, "index": 6000.0}, 6300.0) == pytest.approx(63.0)
    assert M.scale_calibrated({}, 6300.0) is None and M.scale_calibrated({"cap_t": 60, "index": 6000}, None) is None


# ====================================================================
# 보정
# ====================================================================
def test_calibrate_rejects_low_coverage_and_keeps_previous(monkeypatch):
    sp = [f"S{i}" for i in range(500)]
    ndx = [f"S{i}" for i in range(50)] + [f"N{i}" for i in range(50)]
    monkeypatch.setattr(M, "fetch_constituents", lambda state: {"sp500": sp, "ndx": ndx})
    # S&P: 400/500 만 수집(80%) → 거부.  나스닥100: 전부 수집 → 채택
    caps = {t: 1e11 for t in sp[:400]}
    caps.update({t: 2e11 for t in ndx})
    monkeypatch.setattr(M, "fetch_market_caps", lambda tickers, workers=8: caps)
    state = {"calib": {"sp500": {"cap_t": 55.0, "index": 6000.0, "date": "2026-09-01"}}}
    out = M.calibrate(state, {"^GSPC": 6500.0, "^NDX": 24000.0})
    assert out["calib"]["sp500"]["date"] == "2026-09-01"                   # 직전 보정 유지
    assert out["calib"]["ndx"]["cap_t"] == pytest.approx(100 * 2e11 / 1e12) # 20조달러
    assert out["calib"]["ndx"]["n"] == 100 and out["calib"]["ndx"]["index"] == 24000.0
    # union = S0..S499 + N0..N49 = 550, 수집 = S0..S399 + N0..N49 = 450 → 82% < 95% → 보정 없음
    assert "us_union" not in out["calib"]


def test_calibrate_union_coverage_rule(monkeypatch):
    sp = [f"S{i}" for i in range(500)]
    ndx = [f"S{i}" for i in range(50)] + [f"N{i}" for i in range(50)]
    monkeypatch.setattr(M, "fetch_constituents", lambda state: {"sp500": sp, "ndx": ndx})
    caps = {t: 1e11 for t in sp}
    caps.update({t: 1e11 for t in ndx})
    monkeypatch.setattr(M, "fetch_market_caps", lambda tickers, workers=8: caps)
    out = M.calibrate({}, {"^GSPC": 6500.0, "^NDX": 24000.0})
    assert out["calib"]["us_union"]["n_total"] == 550 and out["calib"]["us_union"]["cap_t"] == pytest.approx(55.0)
    assert out["calib"]["sp500"]["cap_t"] == pytest.approx(50.0)


# ====================================================================
# 끝까지 — 가벼운 실행 / 실패 시 stale
# ====================================================================
def _wire(monkeypatch, tmp_path, state, kospi=(3200.0, 900), kosdaq=(450.0, 1700), crypto=None, levels=None):
    monkeypatch.setattr(M, "OUTPUT_FILE", str(tmp_path / "mc.json"))
    monkeypatch.setattr(M, "load_state", lambda name, default=None: state)
    saved = {}
    monkeypatch.setattr(M, "save_state", lambda name, data: saved.update(data) or True)
    monkeypatch.setattr(M, "fetch_usd_krw", lambda last=None: (1400.0, "test"))
    monkeypatch.setattr(M, "fetch_naver_total", lambda market: kospi if market == "KOSPI" else kosdaq)
    monkeypatch.setattr(M, "fetch_crypto", lambda: crypto if crypto is not None else {"cap_t": 3.9, "chg_24h_pct": 1.5, "btc_dominance": 58.0})
    monkeypatch.setattr(M, "fetch_index_levels", lambda syms: levels if levels is not None else {"^GSPC": 6500.0, "^NDX": 24000.0})
    return saved


def test_main_light_run_appends_and_scales(monkeypatch, tmp_path):
    state = {"series": {}, "calib": {"sp500": {"cap_t": 55.0, "index": 6250.0, "date": "2026-09-28", "n": 503, "n_total": 503},
                                      "ndx": {"cap_t": 30.0, "index": 24000.0, "date": "2026-09-28", "n": 101, "n_total": 101}}}
    saved = _wire(monkeypatch, tmp_path, state)
    assert M.main([]) == 0
    out = json.load(open(tmp_path / "mc.json", encoding="utf-8"))
    cards = {c["key"]: c for c in out["cards"]}
    assert cards["kospi"]["value"] == 3200.0 and cards["kospi"]["quality"] == "measured"
    assert cards["domestic"]["value"] == pytest.approx(3650.0) and cards["domestic"]["value_usd_t"] == pytest.approx(3650 / 1400, rel=1e-3)
    assert cards["crypto"]["value"] == 3.9 and cards["crypto"]["chg_1d_pct"] == 1.5 and cards["crypto"]["value_krw_t"] == pytest.approx(3.9 * 1400)
    assert cards["sp500"]["value"] == pytest.approx(55.0 * 6500 / 6250) and cards["sp500"]["quality"] == "estimated"
    assert cards["ndx"]["value"] == pytest.approx(30.0) and cards["us_union"]["quality"] == "unavailable"   # 보정 없음
    assert len(saved["series"]["kospi"]) == 1 and out["fx"]["usd_krw"] == 1400.0
    assert [c["key"] for c in out["cards"]] == ["domestic", "kospi", "kosdaq", "crypto", "sp500", "ndx", "us_union"]


def test_main_marks_stale_when_sources_fail(monkeypatch, tmp_path):
    state = {"series": {"kospi": [{"t": "2026-09-28T15:30", "v": 3100.0}], "crypto": [{"t": "2026-09-28T15:30", "v": 3.8}]}, "calib": {}}
    saved = _wire(monkeypatch, tmp_path, state, kospi=(None, 0), kosdaq=(None, 0), crypto={}, levels={})
    assert M.main([]) == 0
    out = json.load(open(tmp_path / "mc.json", encoding="utf-8"))
    cards = {c["key"]: c for c in out["cards"]}
    assert cards["kospi"]["quality"] == "stale" and cards["kospi"]["value"] == 3100.0      # 직전 값 유지
    assert cards["crypto"]["quality"] == "stale"
    assert cards["kosdaq"]["quality"] == "unavailable" and cards["kosdaq"]["value"] is None
    assert len(saved["series"]["kospi"]) == 1                                              # 실패 시 새 점 없음


def test_main_rejects_out_of_range_values(monkeypatch, tmp_path):
    """네이버 응답 단위가 바뀌어 코스피 합계가 3조원으로 나오면 버려야 한다(sanity)."""
    state = {"series": {}, "calib": {}}
    # fetch_naver_total 내부 sanity 를 우회해 값을 넣는 대신, 범위 함수 자체를 검증
    assert not M._in_range("kospi", 3.0) and M._in_range("kospi", 3200.0)
    assert not M._in_range("crypto", 900.0) and M._in_range("crypto", 3.9)
    assert not M._in_range("sp500", 5.0) and M._in_range("sp500", 58.0)
