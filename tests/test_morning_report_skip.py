"""morning_report.py — 같은 날 두 번 실행 시 건너뛰기(텔레그램 중복 방지) 테스트."""

import os
import sys
import json
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest

MR = pytest.importorskip("morning_report")


def test_main_skips_when_today_report_exists(monkeypatch, tmp_path, capsys):
    monkeypatch.setattr(MR, "DOCS", str(tmp_path))
    today = datetime.now(MR.KST).strftime("%Y-%m-%d")
    (tmp_path / "morning_report.json").write_text(json.dumps({"date": today, "generated_at": f"{today} 07:01:00 KST"}), encoding="utf-8")
    monkeypatch.setattr(sys, "argv", ["morning_report.py"])
    assert MR.main() == 0
    assert "[SKIP]" in capsys.readouterr().out


def test_main_does_not_skip_for_old_report(monkeypatch, tmp_path, capsys):
    """어제 자 리포트만 있으면 건너뛰지 않고 생성 단계로 들어간다 (여기서는 섹션 생성 직전까지만 확인)."""
    monkeypatch.setattr(MR, "DOCS", str(tmp_path))
    (tmp_path / "morning_report.json").write_text(json.dumps({"date": "2000-01-01"}), encoding="utf-8")
    monkeypatch.setattr(sys, "argv", ["morning_report.py"])
    called = {}
    def stop(*a, **k):
        called["L"] = True
        raise RuntimeError("stop-here")
    monkeypatch.setattr(MR, "datetime", MR.datetime)          # no-op: 시간은 그대로
    # 첫 데이터 로드 지점에서 멈춰 '건너뛰지 않았음'만 확인
    import builtins
    real_open = builtins.open
    def fake_open(path, *a, **k):
        if str(path).endswith("morning_report.json"):
            return real_open(path, *a, **k)
        called["L"] = True
        raise RuntimeError("stop-here")
    monkeypatch.setattr(builtins, "open", fake_open)
    try:
        MR.main()
    except RuntimeError:
        pass
    assert called.get("L") is True and "[SKIP]" not in capsys.readouterr().out
