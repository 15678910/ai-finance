"""스모크 테스트: 모든 모듈이 import 가능한지 확인.

실행: python -m pytest tests/test_smoke.py -v
또는: python tests/test_smoke.py
"""

import os
import sys
import importlib

# 프로젝트 루트를 path에 추가
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


# 모든 주요 모듈
MODULES = [
    # core/
    "core",
    "core.env_loader",
    "core.telegram",
    "core.yf_helper",
    "core.state_store",
    # 모니터링
    "breaking_news_monitor",
    "overseas_market_monitor",
    "japan_crisis_monitor",
    "bitcoin_standard_monitor",
    # 분석
    "value_screener",
    "dcf_valuator",
    "earnings_reviewer",
    # AutoResearch
    "auto_research_value",
    "auto_research_portfolio",
    "forward_test_evaluator",
    # 코어 분석
    "commentary_engine",
    "generate_dashboard_data",
]


def test_all_modules_importable():
    """모든 모듈 import 가능 여부 확인."""
    failed = []
    for mod in MODULES:
        try:
            importlib.import_module(mod)
        except Exception as e:
            failed.append((mod, str(e)))

    if failed:
        print("\n실패한 모듈:")
        for mod, err in failed:
            print(f"  ✗ {mod}: {err}")
        raise AssertionError(f"{len(failed)}개 모듈 import 실패")

    print(f"\n✓ 모든 {len(MODULES)}개 모듈 import 성공")


def test_core_functions_available():
    """core/ 공개 함수가 import 가능한지."""
    from core import (
        load_env, get_secret,
        send_message, send_document,
        resolve_ticker, fetch_history, fetch_info_safely,
        load_state, save_state, is_recent_alert, mark_alert_sent,
    )

    # 모든 함수가 callable
    assert callable(load_env)
    assert callable(get_secret)
    assert callable(send_message)
    assert callable(send_document)
    assert callable(resolve_ticker)
    assert callable(fetch_history)
    assert callable(fetch_info_safely)
    assert callable(load_state)
    assert callable(save_state)
    assert callable(is_recent_alert)
    assert callable(mark_alert_sent)


def test_state_store_round_trip():
    """state_store: 저장 → 로드 → 검증."""
    from core import save_state, load_state, mark_alert_sent, is_recent_alert
    import tempfile

    with tempfile.TemporaryDirectory() as tmpdir:
        # 알림 기록
        state = {"last_alerts": {}}
        state = mark_alert_sent(state, "test_alert")

        # 저장
        ok = save_state("test_smoke", state, state_dir=tmpdir)
        assert ok

        # 로드
        loaded = load_state("test_smoke", state_dir=tmpdir)
        assert "last_alerts" in loaded
        assert "test_alert" in loaded["last_alerts"]

        # 최근 알림 체크 (방금 발송)
        assert is_recent_alert(loaded, "test_alert", hours=1) is True

        # 존재하지 않는 알림
        assert is_recent_alert(loaded, "nonexistent", hours=1) is False


def test_resolve_ticker_basic():
    """티커 해석 기본 케이스."""
    from core import resolve_ticker

    # 미국 주식 (변경 없음)
    assert resolve_ticker("AAPL") == "AAPL"
    # 암호화폐 (변경 없음)
    assert resolve_ticker("BTC-USD") == "BTC-USD"
    # 이미 접미사 있음 (변경 없음)
    assert resolve_ticker("005930.KS") == "005930.KS"
    # 빈 문자열
    assert resolve_ticker("") == ""


def test_env_loader_missing_file():
    """존재하지 않는 .env 파일 처리."""
    from core import load_env

    result = load_env("/nonexistent/path/.env")
    assert result == {}


def test_env_loader_parses_quotes():
    """따옴표 처리."""
    from core import load_env
    import tempfile

    with tempfile.NamedTemporaryFile(mode="w", suffix=".env", delete=False, encoding="utf-8") as f:
        f.write('# 주석\n')
        f.write('KEY1=value1\n')
        f.write('KEY2="quoted value"\n')
        f.write("KEY3='single quoted'\n")
        f.write('KEY4=value4  # 인라인 주석\n')
        path = f.name

    try:
        result = load_env(path)
        assert result["KEY1"] == "value1"
        assert result["KEY2"] == "quoted value"
        assert result["KEY3"] == "single quoted"
        assert result["KEY4"] == "value4"
    finally:
        os.unlink(path)


if __name__ == "__main__":
    # pytest 없이 실행
    import traceback

    tests = [
        test_all_modules_importable,
        test_core_functions_available,
        test_state_store_round_trip,
        test_resolve_ticker_basic,
        test_env_loader_missing_file,
        test_env_loader_parses_quotes,
    ]

    passed = 0
    failed = 0
    for test in tests:
        try:
            test()
            print(f"  ✓ {test.__name__}")
            passed += 1
        except Exception as e:
            print(f"  ✗ {test.__name__}: {e}")
            traceback.print_exc()
            failed += 1

    print(f"\n결과: {passed} 통과, {failed} 실패")
    sys.exit(0 if failed == 0 else 1)
