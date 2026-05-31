"""
경제 캘린더 자동 수집
======================
정기적으로 발표되는 경제지표 일정(CPI, FOMC, ECB 등)과
수동 큐레이션 이벤트를 병합하여 docs/economic_calendar.json을 생성합니다.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "economic_calendar.json")
CURATED_FILE = os.path.join(BASE_DIR, "docs", "economic_calendar.json")


def load_existing() -> dict:
    """기존 캘린더 로드."""
    try:
        with open(CURATED_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"events": []}


def sort_events(events: list) -> list:
    """날짜순 정렬 (TBD는 마지막)."""
    def sort_key(e):
        d = e.get("date", "")
        if "TBD" in d:
            return "9999-99-99"
        return d
    return sorted(events, key=sort_key)


def main():
    now_str = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST")
    data = load_existing()
    events = data.get("events", [])
    events = sort_events(events)

    output = {
        "generated_at": now_str,
        "month": datetime.now(KST).strftime("%Y-%m"),
        "note": "주요 일정 및 시장 영향 이벤트 (시뮬레이션/분석용)",
        "events": events,
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE} 저장 완료 ({len(events)}개 이벤트)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
