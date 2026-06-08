"""
주요 이벤트 캘린더 — 텔레그램 알람
==================================
① 아침 다이제스트 (KST 07:00 실행분): 오늘 HIGH 일정 요약
② 임박 알람: HIGH 이벤트 발표 ~1시간 전 (시간별 실행, 이벤트당 1회)
- docs/economic_calendar.json 읽음, 상태는 docs/calendar_alert_state.json (영속)
- 봇: TELEGRAM_FINANCE_BOT_TOKEN / CHAT_ID (환경변수)
"""

import json
import os
import re
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAL_FILE = os.path.join(BASE_DIR, "docs", "economic_calendar.json")
STATE_FILE = os.path.join(BASE_DIR, "docs", "calendar_alert_state.json")

sys.path.insert(0, BASE_DIR)
from core import send_message  # noqa: E402

CAT_EMOJI = {
    "경제지표": "📊", "중앙은행": "🏦", "국채입찰": "🏦", "실적": "📈",
    "파생만기": "🎭", "IPO": "🚀", "지정학": "🌍", "VIP방한": "✈️", "FX/금리": "💱",
}
IMMINENT_MIN = 75  # 이 분 이내면 '임박'으로 발송


def _load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"sent": {}}


def _save_state(state):
    try:
        os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[WARN] 상태 저장 실패: {e}")


def _event_dt(ev):
    """이벤트 발표 시각 → KST datetime (시간 없으면 None)."""
    d = ev.get("date", "")
    m = re.search(r"(\d{1,2}):(\d{2})", ev.get("time", "") or "")
    if not d or not m:
        return None
    hh, mm = m.group(1).zfill(2), m.group(2)
    tz = "-04:00" if re.search(r"ET|EST|EDT", ev.get("time", ""), re.I) else "+09:00"
    try:
        dt = datetime.fromisoformat(f"{d}T{hh}:{mm}:00{tz}")
        return dt.astimezone(KST)
    except Exception:
        return None


def _fmt_event(ev, with_emoji=True):
    ce = CAT_EMOJI.get(ev.get("category"), "•") if with_emoji else ""
    extra = []
    if ev.get("time"):
        extra.append(ev["time"])
    if ev.get("consensus"):
        extra.append(f"컨센 {ev['consensus']}")
    suf = f" ({' · '.join(extra)})" if extra else ""
    return f"🔴{ce} {ev.get('title','')}{suf}"


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    today = now.date().isoformat()
    print(f"[캘린더 알람] {now:%Y-%m-%d %H:%M} KST")

    try:
        with open(CAL_FILE, encoding="utf-8") as f:
            cal = json.load(f)
    except Exception as e:
        print(f"[ERROR] 캘린더 로드 실패: {e}")
        return 1

    # 오늘 HIGH (발표완료 제외)
    high_today = [e for e in cal.get("events", [])
                  if e.get("date") == today and e.get("impact") == "HIGH"
                  and e.get("status") != "released"]

    state = _load_state()
    sent = state.setdefault("sent", {})
    changed = False

    # ── ① 아침 다이제스트 (07시 실행분) ──
    if now.hour == 7:
        dkey = f"digest:{today}"
        if dkey not in sent and high_today:
            ordered = sorted(high_today, key=lambda e: (e.get("time", "")))
            lines = [f"☀️ 오늘({now.month}/{now.day}) 주요 일정 (HIGH)", ""]
            lines += [f"  {_fmt_event(e)}" for e in ordered]
            lines += ["", "📊 실시간 카운트다운: 대시보드 캘린더", "⚠️ 시뮬레이션·분석용."]
            if send_message("\n".join(lines)):
                sent[dkey] = now.isoformat()
                changed = True
                print(f"[OK] 아침 다이제스트 발송 ({len(ordered)}건)")
            else:
                print("[WARN] 다이제스트 발송 실패")

    # ── ② 임박 알람 (발표 ~1시간 전) ──
    for e in high_today:
        dt = _event_dt(e)
        if not dt:
            continue
        mins = (dt - now).total_seconds() / 60
        if 0 < mins <= IMMINENT_MIN:
            ikey = f"imminent:{today}:{e.get('title','')}"
            if ikey in sent:
                continue
            lines = [f"⏰ {int(round(mins))}분 후 발표 임박!", "", _fmt_event(e)]
            if e.get("impact_analysis"):
                lines += ["", e["impact_analysis"][:300]]
            lines += ["", "⚠️ 시뮬레이션·분석용."]
            if send_message("\n".join(lines)):
                sent[ikey] = now.isoformat()
                changed = True
                print(f"[OK] 임박 알람 발송: {e.get('title')} ({int(mins)}분 전)")
            else:
                print(f"[WARN] 임박 알람 발송 실패: {e.get('title')}")

    # 오래된 상태 정리 (7일 경과)
    cutoff = (now - timedelta(days=7)).isoformat()
    for k in [k for k, v in sent.items() if v < cutoff]:
        del sent[k]
        changed = True

    if changed:
        _save_state(state)
    else:
        print("발송할 알람 없음")
    return 0


if __name__ == "__main__":
    sys.exit(main())
