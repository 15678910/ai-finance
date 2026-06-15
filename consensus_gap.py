"""
컨센서스 vs 실데이터 괴리 지표 (한 줄)
========================================
"국내 담론은 낙관(KOSPI 과열·고점)인데, 실데이터는 위험 신호(엔캐리 압력·BOJ 인상·엔화)"
— 이 둘이 동시에 높을 때 '괴리'가 커진다(시장이 위험을 무시 → 현실화 시 반응 지연).

측정 가능한 실데이터만 사용(날조 금지):
  · 시장 안주(complacency): overheating.json avg_heat (과열=도취·안주 대용)
  · 실데이터 위험(risk)   : japan_crisis.json 엔캐리 압력 + USD/JPY + economic_calendar BOJ 인상 확률
'국내 무관심도'는 직접 측정 지표가 없어 숫자화하지 않고 맥락(note)으로만 표기.

출력: docs/consensus_gap.json
🚨 통계·지표 종합. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS, "consensus_gap.json")


def _load(name):
    try:
        with open(os.path.join(DOCS, name), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    jc = _load("japan_crisis.json") or {}
    cal = _load("economic_calendar.json") or {}
    oh = _load("overheating.json") or {}

    # 1) 시장 안주(complacency) — KOSPI/지수 과열도
    complacency = oh.get("avg_heat")

    # 2) 엔캐리 압력 + USD/JPY
    cp = jc.get("carry_pressure") or {}
    carry_score = cp.get("score")
    carry_level = cp.get("level")
    usdjpy = ((jc.get("market") or {}).get("usd_jpy") or {}).get("current")

    # 3) BOJ 인상 확률 + D-day (가장 가까운 미래 이벤트)
    boj_prob = boj_date = boj_dday = None
    events = cal.get("events") if isinstance(cal, dict) else cal
    today = datetime.now(KST).date()
    if isinstance(events, list):
        cands = []
        for e in events:
            if e.get("boj_hike_probability") is not None and e.get("date"):
                try:
                    ed = date.fromisoformat(e["date"])
                except Exception:
                    continue
                if ed >= today:
                    cands.append((ed, e["boj_hike_probability"]))
        if cands:
            cands.sort(key=lambda x: x[0])
            ed, boj_prob = cands[0]
            boj_date = ed.isoformat()
            boj_dday = (ed - today).days

    # ── 위험 점수(0-100): 엔캐리 압력 기반 + 엔화/ BOJ 가산 ──
    risk = float(carry_score) if carry_score is not None else 0.0
    if usdjpy is not None and usdjpy >= 160:
        risk = min(100.0, risk + 8)
    if boj_prob is not None and boj_dday is not None and boj_dday <= 7 and boj_prob >= 80:
        risk = min(100.0, risk + 12)
    risk = round(risk, 1)

    comp = float(complacency) if complacency is not None else 0.0

    # ── 괴리 점수 = 위험 × 안주 (둘 다 높을 때만 큼) ──
    gap_score = round(risk * comp / 100)
    if gap_score >= 45:
        level, color = "큼", "red"
    elif gap_score >= 30:
        level, color = "보통", "amber"
    else:
        level, color = "작음", "green"

    # 한 줄 헤드라인
    risk_bits = []
    if carry_score is not None:
        risk_bits.append(f"엔캐리 압력 {int(carry_score)}{('·'+carry_level) if carry_level else ''}")
    if boj_prob is not None:
        dd = ""
        if boj_dday is not None:
            dd = "D-day" if boj_dday == 0 else (f"D-{boj_dday}" if boj_dday > 0 else "")
        risk_bits.append(f"BOJ 인상 {int(boj_prob)}%{(' '+dd) if dd else ''}")
    if usdjpy is not None:
        risk_bits.append(f"USD/JPY {usdjpy:g}")
    comp_bit = f"KOSPI 과열 {comp:.0f}" if complacency is not None else "시장 과열 데이터 없음"
    headline = f"시장 안주({comp_bit}) ↔ 실데이터 위험({' · '.join(risk_bits) if risk_bits else '데이터 부족'})"

    note = ("국내 담론은 BOJ·엔캐리 청산을 거의 다루지 않는데(낙관·고점), 실데이터(엔캐리 압력·BOJ 인상 확률·엔화)는 "
            "위험을 가리킵니다. 둘이 동시에 높을수록 '괴리'가 커지며, 위험 현실화 시 반응이 늦을 수 있습니다. "
            "BOJ 확률은 수기 컨센서스(실시간 OIS 아님). 지표 종합이며 투자 결정 단독 사용 금지.")

    out = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "gap_level": level, "gap_score": gap_score, "color": color,
        "headline": headline,
        "components": {
            "complacency": {"label": "시장 안주(KOSPI 과열)", "value": comp if complacency is not None else None},
            "carry": {"label": "엔캐리 압력", "value": carry_score, "level": carry_level},
            "boj": {"label": "BOJ 인상 확률", "value": boj_prob, "date": boj_date, "dday": boj_dday},
            "usdjpy": usdjpy,
            "risk_score": risk,
        },
        "note": note,
    }
    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"괴리: {level} ({gap_score}) — {headline}")
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
