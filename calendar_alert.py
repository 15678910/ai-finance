"""
주요 이벤트 캘린더 — 텔레그램 알람
==================================
매일 아침(KST) 오늘·내일의 HIGH/MEDIUM 경제 일정을 텔레그램으로 발송.
- docs/economic_calendar.json 읽음
- 하루 1회 다이제스트 (중복 방지: 날짜 키)
- 봇: TELEGRAM_FINANCE_BOT_TOKEN / CHAT_ID (환경변수)
"""

import json
import os
import sys
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAL_FILE = os.path.join(BASE_DIR, "docs", "economic_calendar.json")

sys.path.insert(0, BASE_DIR)
from core import send_message, load_state, save_state, is_recent_alert, mark_alert_sent  # noqa: E402

IMPACT_EMOJI = {"HIGH": "🔴", "MEDIUM": "🟡", "LOW": "🟢"}
CAT_EMOJI = {
    "경제지표": "📊", "중앙은행": "🏦", "국채입찰": "🏦", "실적": "📈",
    "파생만기": "🎭", "IPO": "🚀", "지정학": "🌍", "VIP방한": "✈️", "FX/금리": "💱",
}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    today = now.date()
    tomorrow = today + timedelta(days=1)
    print(f"[캘린더 알람] {now:%Y-%m-%d %H:%M} KST")

    try:
        with open(CAL_FILE, encoding="utf-8") as f:
            cal = json.load(f)
    except Exception as e:
        print(f"[ERROR] 캘린더 로드 실패: {e}")
        return 1

    # 오늘·내일, HIGH/MEDIUM 이벤트
    targets = []
    for ev in cal.get("events", []):
        d = ev.get("date", "")
        if d not in (today.isoformat(), tomorrow.isoformat()):
            continue
        if ev.get("impact") not in ("HIGH", "MEDIUM"):
            continue
        if ev.get("status") == "released":
            continue
        targets.append(ev)

    if not targets:
        print("발송할 오늘·내일 HIGH/MEDIUM 일정 없음")
        return 0

    # 중복 방지 — 하루 1회
    state = load_state("calendar_alert", {})
    digest_key = f"digest:{today.isoformat()}"
    if is_recent_alert(state, digest_key, hours=18):
        print("오늘 이미 발송함 — 스킵")
        return 0

    targets.sort(key=lambda e: (e.get("date", ""), e.get("time", "")))

    def fmt_day(d):
        return "🗓️ 오늘" if d == today.isoformat() else "🌅 내일"

    lines = [f"📅 주요 경제 일정 알림 ({today.month}/{today.day} 기준)", ""]
    cur_day = None
    for ev in targets:
        d = ev.get("date", "")
        if d != cur_day:
            cur_day = d
            md = d[5:].replace("-", "/")
            lines.append(f"{fmt_day(d)} ({md})")
        ie = IMPACT_EMOJI.get(ev.get("impact"), "")
        ce = CAT_EMOJI.get(ev.get("category"), "•")
        tm = ev.get("time", "")
        cons = ev.get("consensus", "")
        title = ev.get("title", "")
        extra = []
        if tm:
            extra.append(tm)
        if cons:
            extra.append(f"컨센 {cons}")
        suffix = f" ({' · '.join(extra)})" if extra else ""
        lines.append(f"  {ie}{ce} {title}{suffix}")
    lines.append("")
    lines.append("📊 상세·실시간 카운트다운: 대시보드 캘린더 참고")
    lines.append("⚠️ 시뮬레이션·분석용. 투자 결정 단독 사용 금지.")

    text = "\n".join(lines)
    print("--- 발송 내용 ---")
    print(text)

    ok = send_message(text)
    if ok:
        state = mark_alert_sent(state, digest_key)
        save_state("calendar_alert", state)
        print(f"\n[OK] 텔레그램 발송 완료 ({len(targets)}건)")
    else:
        print("\n[WARN] 텔레그램 발송 실패 (토큰/chat_id 확인)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
