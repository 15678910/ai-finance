"""
이벤트 영향권 판정 — 예측 신뢰도 경고용
==========================================
점 예측(개장 전 선물 기반)은 FOMC·BOJ 같은 이벤트 날의 디커플링·안도/실망 랠리를
구조적으로 못 잡는다(백테스트로 신호추가가 MAE를 악화시킴을 확인). 따라서 점 예측을
억지로 고치지 않고, '이 날은 신뢰도가 낮다'를 명시하는 게 정직한 개선이다.

economic_calendar.json에서 target 거래일 ±N일 내 HIGH 중앙은행 이벤트를 찾아 반환.
"""

import json
import os
from datetime import date

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAL_FILE = os.path.join(BASE_DIR, "docs", "economic_calendar.json")
_CB_TAGS = {"FOMC", "Fed", "BOJ", "ECB", "한국은행", "BOK"}


def nearby_cb_events(target_date: str, window_days: int = 1):
    """target_date(YYYY-MM-DD) 기준 ±window_days 내 HIGH 중앙은행 이벤트 목록.
    반환: [{"date","title","dday"}] (dday = 이벤트일 − target, 음수=이미 지남)."""
    try:
        cal = json.load(open(CAL_FILE, encoding="utf-8"))
    except Exception:
        return []
    events = cal.get("events") if isinstance(cal, dict) else cal
    if not isinstance(events, list):
        return []
    try:
        td = date.fromisoformat(target_date)
    except Exception:
        return []
    out = []
    for e in events:
        if e.get("impact") != "HIGH":
            continue
        cat = e.get("category", "")
        tags = e.get("tags") or []
        if cat != "중앙은행" and not (_CB_TAGS & set(tags)):
            continue
        try:
            ed = date.fromisoformat(e.get("date", ""))
        except Exception:
            continue
        dd = (ed - td).days
        if abs(dd) <= window_days:
            out.append({"date": e["date"], "title": e.get("title", ""), "dday": dd})
    out.sort(key=lambda x: abs(x["dday"]))
    return out


def event_risk_block(target_date: str, window_days: int = 1):
    """today 블록에 붙일 이벤트 신뢰도 경고 dict."""
    evs = nearby_cb_events(target_date, window_days)
    return {
        "flag": bool(evs),
        "events": evs,
        "note": ("FOMC·BOJ 등 HIGH 이벤트 ±1거래일 — 개장 전 선물 선행신호의 신뢰도가 낮습니다"
                 "(디커플링·안도/실망 랠리로 방향까지 빗나갈 수 있음). 변동 확대 주의."
                 if evs else None),
    }
