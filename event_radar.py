"""
이벤트 사전영향 레이더 — '내일/이번주 이벤트가 오늘 포지셔닝에 주는 영향'
========================================================================
예측은 밤사이 선물만 보는 게 아니라, 다가오는 이벤트(FOMC·BOJ·PCE·만기 등)가
오늘의 포지셔닝에 주는 영향을 함께 봐야 한다는 관점의 보완 레이어.

핵심: 시장은 이벤트를 '미리' 반영한다 — 임박한 바이너리 이벤트 전엔 관망·변동성 압축,
텔레그래프된(컨센서스 확실) 이벤트는 통과 시 안도 랠리 경향.
※ 정량 예측이 아니라 '이벤트 지형 해석'. economic_calendar.json의 실데이터만 사용.

출력: docs/event_radar.json
"""

import json
import os
import re
import sys
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAL_FILE = os.path.join(BASE_DIR, "docs", "economic_calendar.json")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "event_radar.json")
HORIZON = 7
_CB = {"FOMC", "Fed", "BOJ", "ECB", "BOK", "한국은행"}


def _first_clause(text):
    if not isinstance(text, str):
        return ""
    # 첫 문장/구 (이모지 머리 제거 후 ~90자)
    t = re.split(r"(?<=[.。])\s|·|\n", text.strip())
    s = (t[0] if t else text).strip()
    return (s[:90] + "…") if len(s) > 92 else s


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    try:
        cal = json.load(open(CAL_FILE, encoding="utf-8"))
    except Exception as e:
        print(f"[ERROR] economic_calendar 로드 실패: {e}")
        return 1
    events = cal.get("events") if isinstance(cal, dict) else cal
    if not isinstance(events, list):
        print("[ERROR] events 없음")
        return 1

    today = datetime.now(KST).date()
    radar, seen = [], set()
    for e in sorted(events, key=lambda x: x.get("date", "")):
        if e.get("impact") != "HIGH":
            continue
        try:
            ed = date.fromisoformat(e.get("date", ""))
        except Exception:
            continue
        dd = (ed - today).days
        if not (0 <= dd <= HORIZON):
            continue
        cat = e.get("category", "")
        tags = set(e.get("tags") or [])
        is_cb = (cat == "중앙은행") or bool(_CB & tags)
        key = (e.get("date"), e.get("title", "")[:10])
        if key in seen:
            continue
        seen.add(key)
        cons = None
        if e.get("boj_hike_probability") is not None:
            cons = f"인상확률 {int(e['boj_hike_probability'])}%"
        elif e.get("consensus"):
            cons = str(e["consensus"])
        radar.append({
            "date": e["date"], "dday": dd, "title": e.get("title", ""),
            "category": cat, "region": e.get("region", ""),
            "is_cb": is_cb, "consensus": cons,
            "effect": _first_clause(e.get("impact_analysis") or e.get("detail")),
        })

    # ── 이벤트 지형 바이어스(해석, 정량 예측 아님) ──
    imminent_cb = [r for r in radar if r["is_cb"] and r["dday"] <= 2]
    if imminent_cb:
        nm = imminent_cb[0]["title"]
        bias_key = "event_wait"
        bias = (f"⚠️ 임박 이벤트 대기({nm}, D-{imminent_cb[0]['dday'] or 'day'}) — 결과 전 변동성 압축·관망 우위, "
                "통과 직후 방향 급변 가능. 텔레그래프된 결과면 안도 랠리·실망 시 되돌림.")
    elif any(r["dday"] <= 7 for r in radar):
        bias_key = "approach"
        bias = "HIGH 이벤트 접근(D-3~7) — 점진적 포지셔닝 구간. 선물 신호 신뢰도 보통."
    else:
        bias_key = "calm"
        bias = "임박 HIGH 이벤트 없음 — 펀더멘털·수급 주도, 선물 선행신호 신뢰도 상대적 양호."

    out = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "horizon_days": HORIZON,
        "bias": {"key": bias_key, "text": bias},
        "events": radar,
        "note": ("다가오는 이벤트가 오늘의 포지셔닝에 주는 영향(해석). 시장은 이벤트를 미리 반영하므로 "
                 "점 예측(선물 기반)과 함께 읽어야 함. 정량 예측 아님."),
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"이벤트 레이더: {bias_key} · {len(radar)}건")
    for r in radar:
        print(f"  D+{r['dday']} {r['title'][:30]}{(' ['+r['consensus']+']') if r['consensus'] else ''}")
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
