"""freshness_monitor.py — 정체·내용 이상·주말 유예 판정 테스트."""

import os
import sys
from datetime import datetime, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

FM = pytest.importorskip("freshness_monitor")


def _ts(now, hours_ago):
    return (now - timedelta(hours=hours_ago)).strftime("%Y-%m-%d %H:%M:%S KST")


def test_stale_by_age_and_fresh_ok():
    now = datetime(2026, 11, 3, 10, 0, tzinfo=FM.KST)       # 화요일
    files = [("data.json", 8), ("earnings.json", 30)]
    data = {"data.json": {"generated_at": _ts(now, 9)}, "earnings.json": {"generated_at": _ts(now, 5)}}
    stale, ok, missing, empty = FM.evaluate(files, now, lambda f: data.get(f))
    assert [s["file"] for s in stale] == ["data.json"] and "경과" in stale[0]["reason"]
    assert [o["file"] for o in ok] == ["earnings.json"] and not missing and not empty


def test_fresh_but_empty_is_flagged():
    """2026-10-30~11-02: investor_flow.json 은 매일 새로 생성됐지만 분석 0종목 — 이제 잡힌다."""
    now = datetime(2026, 11, 3, 10, 0, tzinfo=FM.KST)
    data = {"investor_flow.json": {"generated_at": _ts(now, 1), "analyzed_count": 0},
            "market_investor_breakdown.json": {"generated_at": _ts(now, 1), "stale": True, "markets": {"kospi": {"rows": []}}},
            "market_breadth.json": {"generated_at": _ts(now, 1), "breadth": {"n": 797}}}
    stale, ok, missing, empty = FM.evaluate([(f, 30) for f in data], now, lambda f: data.get(f))
    assert {s["file"] for s in stale} == {"investor_flow.json", "market_investor_breakdown.json"}
    assert all("내용 이상" in s["reason"] for s in stale) and len(empty) == 2
    assert [o["file"] for o in ok] == ["market_breadth.json"]


def test_morning_report_must_be_today_after_9am():
    tue_10 = datetime(2026, 11, 3, 10, 0, tzinfo=FM.KST)
    tue_7 = datetime(2026, 11, 3, 7, 0, tzinfo=FM.KST)
    yesterday = {"generated_at": "2026-11-02 08:47:53 KST", "date": "2026-11-02"}
    assert FM.evaluate([("morning_report.json", 27)], tue_10, lambda f: yesterday)[0]        # 10시엔 이상
    assert not FM.evaluate([("morning_report.json", 27)], tue_7, lambda f: yesterday)[0]     # 7시엔 아직 정상
    sat_10 = datetime(2026, 11, 7, 10, 0, tzinfo=FM.KST)
    fri = {"generated_at": "2026-11-06 07:30:00 KST", "date": "2026-11-06"}
    assert not FM.evaluate([("morning_report.json", 27)], sat_10, lambda f: fri)[0]          # 주말엔 금요일 자 허용


def test_weekend_allowance_for_weekday_only_files():
    sun = datetime(2026, 11, 8, 12, 0, tzinfo=FM.KST)
    fri_eve = {"generated_at": "2026-11-06 18:45:00 KST", "breadth": {"n": 800}}
    stale, ok, _, _ = FM.evaluate([("market_breadth.json", 30)], sun, lambda f: fri_eve)
    assert not stale and ok[0]["max_h"] >= 30 + 48
    mon_21 = datetime(2026, 11, 9, 21, 0, tzinfo=FM.KST)                                     # 월요일 밤: 유예 끝
    assert FM.evaluate([("market_breadth.json", 30)], mon_21, lambda f: fri_eve)[0]
    assert FM.weekend_allowance_h(datetime(2026, 11, 4, 12, 0, tzinfo=FM.KST)) == 0          # 수요일


def test_market_cap_and_machine_economy_checks():
    now = datetime(2026, 11, 3, 10, 0, tzinfo=FM.KST)
    mc_bad = {"generated_at": _ts(now, 1), "cards": [{"key": "kospi", "quality": "unavailable"}, {"key": "crypto", "quality": "stale"}]}
    mc_ok = {"generated_at": _ts(now, 1), "cards": [{"key": "kospi", "quality": "measured"}, {"key": "sp500", "quality": "unavailable"}]}
    assert FM.evaluate([("market_cap_tracker.json", 8)], now, lambda f: mc_bad)[0]
    assert not FM.evaluate([("market_cap_tracker.json", 8)], now, lambda f: mc_ok)[0]
    me_bad = {"generated_at": _ts(now, 1), "facts": {"stale": ["stablecoins", "compute"]}}
    assert FM.evaluate([("machine_economy.json", 8)], now, lambda f: me_bad)[0]


def test_missing_file_listed():
    now = datetime(2026, 11, 3, 10, 0, tzinfo=FM.KST)
    _, _, missing, _ = FM.evaluate([("nope.json", 8)], now, lambda f: None)
    assert missing == ["nope.json"]
