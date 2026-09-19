"""portfolio_quotes.py 파서·폴백 테스트.

네이버는 이 환경에서 접근이 막혀 있어 네트워크 경로는 모킹한다. 파서는 순수 함수라
응답 형태만 고정하면 검증할 수 있다 — 폴링 API 형태는 naver-stock-proxy-worker.js 에
문서화된 것을 그대로 쓴다.

실행: python -m pytest tests/test_portfolio_quotes.py -v
"""

import os
import sys
import json

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

PQ = pytest.importorskip("portfolio_quotes")


# ====================================================================
# 파서
# ====================================================================
def test_parse_polling_json_worker_documented_shape():
    payload = {"datas": [
        {"itemCode": "000660", "stockName": "SK하이닉스", "closePrice": "2,343,000", "fluctuationsRatio": "-2.56"},
        {"itemCode": "005930", "stockName": "삼성전자", "closePrice": "261,000"},
        {"itemCode": "bad", "stockName": "x", "closePrice": "1"},          # 코드 형식 불량
        {"itemCode": "123456", "stockName": "y", "closePrice": "-"},        # 가격 없음
    ]}
    q = PQ.parse_polling_json(payload)
    assert q == {"000660": {"n": "SK하이닉스", "p": 2343000},
                 "005930": {"n": "삼성전자", "p": 261000}}


def test_parse_ranking_json_accepts_variants():
    # 최상위 리스트
    a = PQ.parse_ranking_json([{"itemCode": "005930", "stockName": "삼성전자", "closePrice": "261,000"}])
    # {"stocks":[...]} 래핑 + 대체 필드명
    b = PQ.parse_ranking_json({"stocks": [{"cd": "005930", "nm": "삼성전자", "nowVal": "261000"}]})
    # 이중 래핑
    c = PQ.parse_ranking_json({"result": {"stocks": [{"code": "005930", "name": "삼성전자", "price": 261000}]}})
    assert a == b == c == {"005930": {"n": "삼성전자", "p": 261000}}


def test_parse_ranking_json_rejects_garbage():
    assert PQ.parse_ranking_json(None) == {}
    assert PQ.parse_ranking_json({"error": "blocked"}) == {}
    assert PQ.parse_ranking_json("<html>차단</html>") == {}
    assert PQ.parse_ranking_json([{"itemCode": "005930"}]) == {}          # 가격 없음


def test_parse_html_rows_legacy_regex():
    html = ('<a href="/item/main.naver?code=005930" class="tltle">삼성전자</a></td>'
            '<td class="number">261,000</td>'
            '<a href="/item/main.naver?code=000660" class="tltle">SK하이닉스</a></td>'
            '<td class="number">1,596,000</td>')
    assert PQ.parse_html_rows(html) == {"005930": {"n": "삼성전자", "p": 261000},
                                        "000660": {"n": "SK하이닉스", "p": 1596000}}
    assert PQ.parse_html_rows("<html>구조 변경</html>") == {}


# ====================================================================
# 폴백 순서
# ====================================================================
def _rows(n, start=0):
    return {f"{100000 + start + i:06d}": {"n": f"s{i}", "p": 1000 + i} for i in range(n)}


def test_collect_uses_ranking_api_when_sufficient(monkeypatch):
    monkeypatch.setattr(PQ, "fetch_ranking_api", lambda m: _rows(300, 0 if m == "KOSPI" else 300))
    monkeypatch.setattr(PQ, "fetch_polling_api", lambda c: pytest.fail("폴링까지 가면 안 됨"))
    monkeypatch.setattr(PQ, "scrape_html", lambda s: pytest.fail("HTML까지 가면 안 됨"))
    q, src = PQ.collect()
    assert src == "naver-mobile-api" and len(q) == 600


def test_collect_falls_back_to_polling_with_previous_codes(monkeypatch, tmp_path):
    prev = tmp_path / "quotes.json"
    prev.write_text(json.dumps({"q": _rows(500)}), encoding="utf-8")
    monkeypatch.setattr(PQ, "OUTPUT_FILE", str(prev))
    monkeypatch.setattr(PQ, "fetch_ranking_api", lambda m: {})                 # 1순위 0건
    monkeypatch.setattr(PQ, "fetch_polling_api", lambda codes: _rows(len(codes)))
    monkeypatch.setattr(PQ, "scrape_html", lambda s: pytest.fail("HTML까지 가면 안 됨"))
    q, src = PQ.collect()
    assert src == "naver-polling-api" and len(q) == 500


def test_collect_falls_back_to_html_last(monkeypatch, tmp_path):
    monkeypatch.setattr(PQ, "OUTPUT_FILE", str(tmp_path / "none.json"))    # 직전 파일 없음
    monkeypatch.setattr(PQ, "fetch_ranking_api", lambda m: {})
    monkeypatch.setattr(PQ, "fetch_polling_api", lambda c: pytest.fail("코드 없으면 폴링 생략"))
    monkeypatch.setattr(PQ, "scrape_html", lambda s: _rows(120, 0 if s == 0 else 120))
    q, src = PQ.collect()
    assert src == "naver-html" and len(q) == 240


def test_main_preserves_file_when_all_sources_fail(monkeypatch, tmp_path, capsys):
    """전 경로 0건이면 기존 파일을 건드리지 않고 1을 돌려준다 (예전과 동일)."""
    prev = tmp_path / "quotes.json"
    prev.write_text('{"q":{"005930":{"n":"삼성전자","p":1}}}', encoding="utf-8")
    monkeypatch.setattr(PQ, "OUTPUT_FILE", str(prev))
    monkeypatch.setattr(PQ, "collect", lambda: ({}, "naver-html"))
    assert PQ.main() == 1
    assert json.loads(prev.read_text(encoding="utf-8"))["q"]["005930"]["p"] == 1
    assert "[ERROR] 수집 부족" in capsys.readouterr().out


def test_main_writes_source_field(monkeypatch, tmp_path):
    out = tmp_path / "quotes.json"
    monkeypatch.setattr(PQ, "OUTPUT_FILE", str(out))
    monkeypatch.setattr(PQ, "collect", lambda: (_rows(150), "naver-mobile-api"))
    assert PQ.main() == 0
    d = json.loads(out.read_text(encoding="utf-8"))
    assert d["count"] == 150 and d["source"] == "naver-mobile-api"
    assert set(d) >= {"generated_at", "count", "source", "q", "note"}     # 프론트 계약 유지


def test_diag_logged_on_zero_rows(monkeypatch, capsys):
    """0건이면 응답 앞부분이 로그에 남아야 한다 — 이번 장애는 이 줄이 없어 진단 불가였다."""
    monkeypatch.setattr(PQ, "_get", lambda *a, **k: "<html>Access Denied</html>")
    monkeypatch.setattr(PQ, "HTML_PAGES", 1)
    assert PQ.scrape_html(0) == {}
    assert "[DIAG]" in capsys.readouterr().out
