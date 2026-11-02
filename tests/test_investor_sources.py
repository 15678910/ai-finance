"""투자자 매매 데이터 소스 교체 테스트 — core/naver_flow, investor_flow_monitor 대체 경로, market_investor_breakdown KRX 파서."""

import os
import sys
import json

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

NF = pytest.importorskip("core.naver_flow")
IFM = pytest.importorskip("investor_flow_monitor")
MIB = pytest.importorskip("market_investor_breakdown")


def test_naver_flow_parse_holding_fields():
    rows = NF.parse_trend_json([{"bizdate": "20261102", "closePrice": "71,500", "foreignerPureBuyQuant": "+1,000", "organPureBuyQuant": "-500",
                                 "foreignerHoldRatio": "52.13%", "foreignerHoldQuant": "3,111,000,000"},
                                {"bizdate": "20261030", "closePrice": "70,000", "foreignerPureBuyQuant": "1", "organPureBuyQuant": "1"}])
    assert rows[0]["foreign_holding_pct"] == 52.13 and rows[0]["foreign_holding_shares"] == 3_111_000_000   # '52.13%' 도 읽어야 한다
    # 2026-11-02 실제 응답 키 그대로
    real = [{"accumulatedTradingVolume": "12,345,678", "bizdate": "20261102", "closePrice": "71,500", "compareToPreviousClosePrice": "1,500",
             "compareToPreviousPrice": {"code": "2", "text": "상승"}, "foreignerHoldRatio": "52.13%", "foreignerPureBuyQuant": "+123,456",
             "individualPureBuyQuant": "-100,000", "itemCode": "005930", "organPureBuyQuant": "-23,456"}]
    r = NF.parse_trend_json(real)[0]
    assert r["foreign_holding_pct"] == 52.13 and r["foreign_net_shares"] == 123456 and r["institutional_net_shares"] == -23456 and r["volume"] == 12345678
    assert rows[0]["change_pct"] == pytest.approx(2.14, abs=0.01)
    assert set(rows[0]) >= {"date", "close", "change_pct", "volume", "institutional_net_shares", "foreign_net_shares", "foreign_holding_shares", "foreign_holding_pct"}


def test_investor_flow_monitor_uses_mobile_then_html(monkeypatch):
    calls = []
    monkeypatch.setattr(IFM, "fetch_trend_rows", lambda code, n=20, get=None: calls.append("mobile") or [])
    monkeypatch.setattr(IFM, "fetch_naver_flow_html", lambda code: calls.append("html") or [{"date": "2026-11-02"}])
    assert IFM.fetch_naver_flow("005930") == [{"date": "2026-11-02"}] and calls == ["mobile", "html"]
    calls.clear()
    monkeypatch.setattr(IFM, "fetch_trend_rows", lambda code, n=20, get=None: calls.append("mobile") or [{"date": "2026-11-02", "mobile": True}])
    assert IFM.fetch_naver_flow("005930")[0]["mobile"] is True and calls == ["mobile"]


def _krx_row(d, vals):
    row = {"TRD_DD": d}
    for i, v in enumerate(vals, start=1):
        row[f"TRDVAL{i}"] = f"{v:,}"
    return row


def test_parse_krx_output_maps_columns_and_scales_to_million_won():
    # 금융투자 보험 투신 사모 은행 기타금융 연기금 기관합계 기타법인 개인 외국인 기타외국인 (원)
    inst = [1_000e8, 200e8, -300e8, 50e8, 10e8, 5e8, 2_000e8]
    inst_total = sum(inst)
    other_corp, personal, foreign, other_foreign = 100e8, -4_000e8, 900e8, 34.9e8   # 합 ≈ 0
    vals = [int(x) for x in inst + [inst_total, other_corp, personal, foreign, other_foreign]]
    rows = MIB.parse_krx_output({"output": [_krx_row("2026/11/02", vals), _krx_row("2026/10/30", vals)]})
    assert len(rows) == 2 and rows[0]["date"] == "2026-11-02" and rows[0]["unit"] == "백만원"
    assert rows[0]["financial_inv"] == 100_000 and rows[0]["pension"] == 200_000        # 1,000억 = 100,000백만원
    assert rows[0]["foreign"] == int(round((900e8 + 34.9e8) / 1e6)) and rows[0]["institutional_total"] == int(round(inst_total / 1e6))
    assert MIB.validate_breakdown(rows) is True
    assert MIB.parse_krx_output({"output": [{"TRD_DD": "2026/11/02", "TRDVAL1": "1"}]}) == []     # 열 부족 → 버림
    assert MIB.parse_krx_output({"error": 1}) == [] and MIB.parse_krx_output(None) == []


def test_validate_breakdown_rejects_wrong_mapping():
    good = [{"personal": -100, "foreign": 60, "institutional_total": 35, "other_corp": 5}] * 5
    bad = [{"personal": 100, "foreign": 60, "institutional_total": 35, "other_corp": 5}] * 5
    assert MIB.validate_breakdown(good) and not MIB.validate_breakdown(bad) and not MIB.validate_breakdown([])


def test_parse_naver_3class_scales_and_marks_detail_unavailable():
    # 억원 단위 응답 (외국인 중앙값 ~ 수백) → ×100 으로 백만원
    payload = {"result": [{"bizdate": "20261102", "individualPureBuyAmount": "-1,234", "foreignerPureBuyAmount": "+800", "organPureBuyAmount": "400"},
                          {"bizdate": "20261030", "individualPureBuyAmount": "500", "foreignerPureBuyAmount": "-900", "organPureBuyAmount": "350"}]}
    rows = MIB.parse_naver_3class(payload)
    assert rows[0]["date"] == "2026-11-02" and rows[0]["foreign"] == 80_000 and rows[0]["personal"] == -123_400
    assert rows[0]["other_corp"] == 3_400 and rows[0]["detail_available"] is False and rows[0]["pension"] == 0
    assert MIB.validate_breakdown(rows)
    # 백만원 단위 응답 (중앙값 > 50,000) → 그대로
    big = {"result": [{"bizdate": "20261102", "individualNetBuy": "-120,000,000", "foreignNetBuy": "80,000,000", "institutionNetBuy": "35,000,000"}]}
    assert MIB.parse_naver_3class(big)[0]["foreign"] == 80_000_000
    assert MIB.parse_naver_3class({"result": [{"bizdate": "20261102", "closePrice": 1}]}) == []     # 키 없음
    assert MIB.parse_naver_3class("<html>") == []


def test_collect_market_order_krx_mobile_pc(monkeypatch):
    monkeypatch.setattr(MIB, "fetch_krx_breakdown", lambda code, days=30: [])
    monkeypatch.setattr(MIB, "fetch_naver_mobile_3class", lambda code, days=30: [{"date": "2026-11-01", "detail_available": False}])
    assert MIB.collect_market("01")[0]["detail_available"] is False
    monkeypatch.setattr(MIB, "fetch_naver_mobile_3class", lambda code, days=30: [])
    monkeypatch.setattr(MIB, "fetch_breakdown_for_date", lambda code, d: {"2026-10-30": {"date": "2026-10-30"}})
    assert MIB.collect_market("01")[0]["date"] == "2026-10-30"


def test_collect_market_prefers_krx_then_naver(monkeypatch):
    monkeypatch.setattr(MIB, "fetch_naver_mobile_3class", lambda code, days=30: [])
    monkeypatch.setattr(MIB, "fetch_krx_breakdown", lambda code, days=30: [{"date": "2026-11-02", "personal": 1}])
    assert MIB.collect_market("01")[0]["date"] == "2026-11-02"
    monkeypatch.setattr(MIB, "fetch_krx_breakdown", lambda code, days=30: [])
    monkeypatch.setattr(MIB, "fetch_breakdown_for_date", lambda code, d: {"2026-10-30": {"date": "2026-10-30"}})
    assert MIB.collect_market("01")[0]["date"] == "2026-10-30"


def test_breakdown_main_keeps_previous_file_on_total_failure(monkeypatch, tmp_path):
    docs = tmp_path / "docs"; docs.mkdir()
    prev = {"generated_at": "2026-10-30", "markets": {"kospi": {"rows": [{"date": "2026-10-30"}]}}}
    (docs / "market_investor_breakdown.json").write_text(json.dumps(prev), encoding="utf-8")
    monkeypatch.setattr(MIB.os.path, "abspath", lambda p: str(tmp_path / "x.py"))
    monkeypatch.setattr(MIB, "collect_market", lambda code, days_target=30: [])
    monkeypatch.setattr(MIB.time, "sleep", lambda s: None)
    assert MIB.main() == 1
    kept = json.loads((docs / "market_investor_breakdown.json").read_text(encoding="utf-8"))
    assert kept["stale"] is True and kept["markets"]["kospi"]["rows"][0]["date"] == "2026-10-30"
