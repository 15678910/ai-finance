"""morning_report.breadth_divergence — 지수와 종목 다수가 엇갈린 날에만 경고하는지."""

import os
import sys
import json
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

MR = pytest.importorskip("morning_report")


def _mb(as_of="2026-11-13", kospi=(0.8, 126, 257), kosdaq=(0.1, 200, 190)):
    def side(t):
        return {"date": as_of, "change_pct": t[0]}, {"n": 400, "advancers": t[1], "decliners": t[2]}
    (ki, kb), (qi, qb) = side(kospi), side(kosdaq)
    return {"as_of": as_of, "index": {"KOSPI": ki, "KOSDAQ": qi}, "by_market": {"KOSPI": kb, "KOSDAQ": qb}}


MON = date(2026, 11, 16)


def test_narrow_rally_warns_only_for_diverging_market():
    out = MR.breadth_divergence(_mb(), MON)
    assert len(out) == 1
    assert "KOSPI +0.80%" in out[0] and "상승 126 / 하락 257" in out[0] and "좁은 상승" in out[0]


def test_index_only_decline_is_noted():
    out = MR.breadth_divergence(_mb(kospi=(-0.9, 280, 110)), MON)
    assert len(out) == 1 and "KOSPI -0.90%" in out[0] and "버텼다" in out[0]


def test_no_warning_when_aligned_or_small_move():
    assert MR.breadth_divergence(_mb(kospi=(1.2, 300, 90)), MON) == []           # 넓은 상승
    assert MR.breadth_divergence(_mb(kospi=(0.4, 100, 290)), MON) == []          # 0.5% 미만 — 잡음
    assert MR.breadth_divergence(_mb(kospi=(0.5, 160, 240)), MON) == []          # 비율 40% = 경계, 경고 아님


def test_threshold_edges():
    assert len(MR.breadth_divergence(_mb(kospi=(0.5, 159, 241)), MON)) == 1      # 39.75% < 40
    assert len(MR.breadth_divergence(_mb(kospi=(-0.5, 241, 159)), MON)) == 1     # 60.25% > 60


def test_stale_future_or_missing_data_is_silent():
    assert MR.breadth_divergence(_mb(as_of="2026-11-10"), MON) == []             # 6일 전
    assert MR.breadth_divergence(_mb(as_of="2026-11-17"), MON) == []             # 미래 날짜
    assert MR.breadth_divergence({}, MON) == []
    old = _mb()
    del old["index"]["KOSPI"]["change_pct"]                                       # 수집기 갱신 전 JSON
    assert MR.breadth_divergence(old, MON) == []
    zero = _mb(kospi=(0.8, 0, 0))
    assert MR.breadth_divergence(zero, MON) == []
    bad = _mb()
    bad["by_market"]["KOSPI"]["advancers"] = None
    assert MR.breadth_divergence(bad, MON) == []


def test_index_bar_from_other_day_is_not_compared():
    """지수는 종목과 따로 받는다 — 지수 마지막 봉이 종목 집계일과 다르면 괴리를 판단하지 않는다."""
    mb = _mb()
    mb["index"]["KOSPI"]["date"] = "2026-11-12"
    assert MR.breadth_divergence(mb, MON) == []
    del mb["index"]["KOSPI"]["date"]                                              # 날짜 필드 없는 옛 JSON
    assert MR.breadth_divergence(mb, MON) == []


def test_main_adds_section_and_skips_telegram(monkeypatch, tmp_path):
    """전체 실행: 섹션이 JSON에 들어가고, 테스트에서는 텔레그램을 보내지 않는다."""
    import core
    today = datetime.now(MR.KST).date()
    (tmp_path / "market_breadth.json").write_text(json.dumps(_mb(as_of=today.isoformat())), encoding="utf-8")
    monkeypatch.setattr(MR, "DOCS", str(tmp_path))
    monkeypatch.setattr(sys, "argv", ["morning_report.py", "--force"])
    monkeypatch.setattr(core, "get_secret", lambda name: None)
    monkeypatch.setattr(core, "send_message", lambda *a, **k: pytest.fail("텔레그램 발송 시도"))
    assert MR.main() == 0
    rep = json.loads((tmp_path / "morning_report.json").read_text(encoding="utf-8"))
    sec = [s for s in rep["sections"] if s["title"].startswith("📊 시장 폭 괴리")]
    assert len(sec) == 1 and any("좁은 상승" in x for x in sec[0]["lines"])
