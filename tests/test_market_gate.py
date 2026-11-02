"""market_breadth.py · flow_streak_screener.py · core/kr_universe.py 테스트 (네트워크 모킹)."""

import os
import sys
import json
from datetime import date, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

B = pytest.importorskip("market_breadth")
F = pytest.importorskip("flow_streak_screener")
U = pytest.importorskip("core.kr_universe")


# ====================================================================
# 유니버스
# ====================================================================
def test_parse_universe_page_excludes_etf_and_keeps_name():
    p = {"stocks": [{"itemCode": "005930", "stockName": "삼성전자", "marketValue": "4,000,000", "stockEndType": "stock"},
                    {"itemCode": "069500", "stockName": "KODEX 200", "marketValue": "100,000", "stockEndType": "etf"},
                    {"itemCode": "000660", "stockName": "SK하이닉스", "marketValue": "2,000,000"}]}
    out = U.parse_universe_page(p)
    assert [o["code"] for o in out] == ["005930", "000660"] and out[0]["name"] == "삼성전자" and out[0]["mv_eok"] == 4_000_000


def test_fetch_universe_falls_back_to_previous_when_short(monkeypatch):
    monkeypatch.setattr(U, "_fetch_market", lambda market, n: [{"code": "005930", "name": "삼성전자", "mv_eok": 1}])
    monkeypatch.setattr(U, "load_state", lambda name, default=None: {"universe": [{"code": "x", "name": "이전", "market": "KOSPI", "mcap_b": 1}]})
    monkeypatch.setattr(U, "save_state", lambda name, data: True)
    assert U.fetch_universe()[0]["name"] == "이전"
    assert U.yf_symbol("005930", "KOSPI") == "005930.KS" and U.yf_symbol("035720", "KOSDAQ") == "035720.KQ"


# ====================================================================
# 시장폭
# ====================================================================
def _ramp(n, start, step):
    return [start + step * i for i in range(n)]


def test_stock_metrics_and_breadth():
    up = _ramp(260, 100, 1)                      # 꾸준히 상승 → 모든 MA 위, 신고가
    down = _ramp(260, 400, -1)                   # 꾸준히 하락 → MA 아래, 신저가
    short = [10, 11, 12]                         # 이력 부족 → MA·52주 None
    m = B.stock_metrics(up)
    assert m["above20"] and m["above50"] and m["above200"] and m["high52"] and not m["low52"] and m["adv"]
    assert B.stock_metrics(short)["above20"] is None and B.stock_metrics(short)["high52"] is None
    b = B.compute_breadth({"a": up, "b": down, "c": short})
    assert b["n"] == 3 and b["eligible_200"] == 2
    assert b["pct_above_50"] == 50.0 and b["new_highs"] == 1 and b["new_lows"] == 1 and b["advancers"] == 2 and b["decliners"] == 1


def test_score_grade_and_gate():
    strong = B.score_breadth({"pct_above_20": 80, "pct_above_50": 75, "pct_above_200": 70, "new_highs": 30, "new_lows": 5})
    assert strong["grade"] == "강세" and strong["gate_open"] is True and strong["score"] == pytest.approx((80 + 75 + 70 + 30 / 35 * 100) / 4, abs=0.1)
    weak = B.score_breadth({"pct_above_20": 20, "pct_above_50": 15, "pct_above_200": 10, "new_highs": 11, "new_lows": 215})
    assert weak["grade"] == "극약세" and weak["gate_open"] is False and "50일선" in weak["reason"]
    mid = B.score_breadth({"pct_above_20": 55, "pct_above_50": 40, "pct_above_200": 45, "new_highs": 0, "new_lows": 0})
    assert mid["hl_pct"] == 50.0 and mid["gate_open"] is True
    assert B.score_breadth({"pct_above_20": None})["gate_open"] is None


def test_index_context_ftd_and_distribution_days():
    # 고점 1000 → 저점 760 (-24%) → 반등. 저점 후 5일차에 +2% & 거래량 증가 → FTD
    closes = _ramp(100, 900, 1) + [1000] + [1000 - 24 * i for i in range(1, 11)]                # 저점 760 (i=10)
    post = [765, 770, 775, 780, 796, 800, 805, 810]                                              # 5일차(780→796 = +2.05%)
    closes += post
    vols = [100.0] * len(closes)
    i_lo = len(closes) - len(post) - 1
    vols[i_lo + 5] = 150.0                                                                       # FTD 거래량 증가
    dates = [(date(2026, 1, 1) + timedelta(days=i)).isoformat() for i in range(len(closes))]
    ctx = B.index_context(closes, vols, dates)
    assert ctx["max_drawdown_pct"] == pytest.approx(-24.0) and ctx["low_date"] == dates[i_lo]
    assert ctx["days_since_low"] == len(post) and ctx["ftd"]["day_n"] == 5 and ctx["ftd"]["gain_pct"] == pytest.approx(2.05, abs=0.01)
    assert ctx["rebound_from_low_pct"] == pytest.approx((810 / 760 - 1) * 100, abs=0.01)
    # 분배일: 마지막 25일 중 하락 & 거래량 증가 — 하락 구간은 거래량 동일(100)이라 0
    assert ctx["distribution_days_25"] == 0
    vols2 = list(vols)
    for d in range(len(closes) - 20, len(closes) - 9):                                           # 하락 구간에 거래량 증가 부여
        vols2[d] = 100.0 + (d % 3)
    ctx2 = B.index_context(closes, vols2, dates)
    assert ctx2["distribution_days_25"] >= 1
    assert B.index_context([1, 2, 3], None, ["a", "b", "c"]) == {}


def test_breadth_main_end_to_end(monkeypatch, tmp_path):
    monkeypatch.setattr(B, "OUTPUT_FILE", str(tmp_path / "mb.json"))
    monkeypatch.setattr(B, "fetch_universe", lambda a, b: [{"code": "005930", "name": "삼성전자", "market": "KOSPI", "mcap_b": 400},
                                                            {"code": "035720", "name": "카카오", "market": "KOSDAQ", "mcap_b": 20}])
    dates = [(date(2025, 10, 1) + timedelta(days=i)).isoformat() for i in range(260)]
    hist = {"005930.KS": {"dates": dates, "close": _ramp(260, 100, 1), "volume": [1.0] * 260},
            "035720.KQ": {"dates": dates, "close": _ramp(260, 400, -1), "volume": [1.0] * 260},
            "^KS11": {"dates": dates, "close": _ramp(260, 3000, 2), "volume": [0.0] * 260},
            "069500.KS": {"dates": dates, "close": _ramp(260, 30000, 20), "volume": [100.0] * 260}}
    monkeypatch.setattr(B, "fetch_price_history", lambda syms, period="1y", chunk=150: {s: hist[s] for s in syms if s in hist})
    saved = {}
    monkeypatch.setattr(B, "load_state", lambda name, default=None: {"history": [{"date": "2026-01-01", "score": 40.0}]})
    monkeypatch.setattr(B, "save_state", lambda name, data: saved.update(data) or True)
    assert B.main() == 0
    out = json.load(open(tmp_path / "mb.json", encoding="utf-8"))
    assert out["breadth"]["n"] == 2 and out["score"] == 50.0 and out["gate"]["open"] is True
    assert out["by_market"]["KOSPI"]["pct_above_50"] == 100.0 and "KOSPI" in out["index"] and "KOSDAQ" not in out["index"]
    assert len(saved["history"]) == 2 and out["history"][-1][0] == dates[-1]


def test_breadth_main_fails_loudly_on_low_coverage(monkeypatch, tmp_path):
    monkeypatch.setattr(B, "OUTPUT_FILE", str(tmp_path / "mb.json"))
    monkeypatch.setattr(B, "fetch_universe", lambda a, b: [{"code": f"{i:06d}", "name": str(i), "market": "KOSPI", "mcap_b": 1} for i in range(10)])
    monkeypatch.setattr(B, "fetch_price_history", lambda syms, period="1y", chunk=150: {})
    assert B.main() == 1 and not (tmp_path / "mb.json").exists()


# ====================================================================
# 수급 연속
# ====================================================================
def _rows(f, i, close=10000):
    return [{"date": f"2026-10-{20 - k:02d}", "close": close, "change_pct": 1.0, "foreign_net_shares": a, "institutional_net_shares": b}
            for k, (a, b) in enumerate(zip(f, i))]


def test_streak_sign_and_length():
    assert F.streak([5, 3, 1, -2, 4]) == 3 and F.streak([-1, -1, 0, -1]) == -2 and F.streak([0, 5]) == 0 and F.streak([]) == 0


def test_compute_streaks():
    s = F.compute_streaks(_rows([100, 200, 300, -50, 10], [50, 60, -5, 70, 80], close=1_000_000))
    assert s["f_streak"] == 3 and s["i_streak"] == 2 and s["joint"] == 2 and s["joint_sell"] == 0
    assert s["f5_eok"] == pytest.approx(5.6, abs=0.01) and s["as_of"] == "2026-10-20"          # 560주 × 100만원 = 5.6억
    assert F.compute_streaks([]) == {}
    # 날짜가 섞여 들어와도 최신순으로 정렬해 계산
    shuffled = list(reversed(_rows([1, 1, -1], [1, 1, -1])))
    assert F.compute_streaks(shuffled)["joint"] == 2


def test_screen_filters_and_orders():
    uni = [{"code": "A", "name": "a", "market": "KOSPI", "mcap_b": 1}, {"code": "B", "name": "b", "market": "KOSDAQ", "mcap_b": 1},
           {"code": "C", "name": "c", "market": "KOSPI", "mcap_b": 1}, {"code": "D", "name": "d", "market": "KOSPI", "mcap_b": 1},
           {"code": "E", "name": "e", "market": "KOSPI", "mcap_b": 1}]
    flows = {"A": _rows([1, 1, 1, 1], [1, 1, 1, 1]),           # 동반 4일
             "B": _rows([1, 1, 1, 1, 1, 1], [-1, -1, -1, -1, -1, -1]),  # 외인 6일, 기관 매도
             "C": _rows([1, 1, -1], [1, 1, -1]),               # 동반 2일 — 3일 미만이라 buy 제외, joint2 포함
             "D": _rows([-1, -1, -1], [-1, -1, -1]),           # 동반 매도 3일
             }                                                   # E: 수집 실패
    r = F.screen(uni, flows)
    assert [x["code"] for x in r["buy"]] == ["A", "B"]           # 동반 연속일 우선, 그다음 최장 연속일
    assert [x["code"] for x in r["joint2"]] == ["A", "C"]
    assert [x["code"] for x in r["sell"]] == ["D"]
    assert r["counts"] == {"universe": 5, "fetched": 4, "buy": 2, "joint2": 2, "sell": 1}


def test_screener_main_end_to_end(monkeypatch, tmp_path):
    monkeypatch.setattr(F, "OUTPUT_FILE", str(tmp_path / "fs.json"))
    bfile = tmp_path / "mb.json"
    bfile.write_text(json.dumps({"gate": {"open": False, "reason": "종합 16점"}, "score": 16.6, "grade": "극약세", "as_of": "2026-10-20"}), encoding="utf-8")
    monkeypatch.setattr(F, "BREADTH_FILE", str(bfile))
    uni = [{"code": "A", "name": "a", "market": "KOSPI", "mcap_b": 1}, {"code": "B", "name": "b", "market": "KOSPI", "mcap_b": 1}]
    monkeypatch.setattr(F, "fetch_universe", lambda a, b: uni)
    monkeypatch.setattr(F, "fetch_naver_flow", lambda code: _rows([1, 1, 1], [1, 1, 1]) if code == "A" else [])
    monkeypatch.setattr(F.time, "sleep", lambda s: None)
    assert F.main() == 0
    out = json.load(open(tmp_path / "fs.json", encoding="utf-8"))
    assert out["gate_blocked"] is True and out["gate"]["grade"] == "극약세"
    assert out["counts"]["buy"] == 1 and out["buy"][0]["code"] == "A" and out["as_of"] == "2026-10-20"
    # 전부 실패 → 덮어쓰지 않고 실패
    monkeypatch.setattr(F, "fetch_naver_flow", lambda code: [])
    os.remove(tmp_path / "fs.json")
    assert F.main() == 1 and not (tmp_path / "fs.json").exists()
