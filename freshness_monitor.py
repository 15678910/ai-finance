"""
데이터 신선도 감시기 — 정체 자동 탐지 + 텔레그램 경보
========================================================
재발 방지: 워크플로가 조용히 실패해(commit 누락 등) 데이터가 멈춰도
며칠간 모르던 사고(2026-06-17~18 market_news 19시간 정체)를 막는다.

각 docs/*.json의 generated_at을 읽어 '예상 최대 나이'를 넘으면 정체로 판정,
텔레그램으로 즉시 경보(중복 방지). 큐레이션/월간 콘텐츠는 감시 제외.

출력: docs/freshness.json (상태 요약) + 정체 시 텔레그램.
"""

import json
import os
import re
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS, "freshness.json")

# (파일, 예상 최대 나이[시간]) — KOSPI 장중엔 overseas-market cron이 쉬어서 ~7h 정상 → 여유 있게.
#   ※ 큐레이션/월간(key_debates·trade_data)은 감시 제외(본래 저빈도).
MONITORED = [
    ("data.json", 8), ("market_news.json", 10), ("kospi_scenario.json", 10),
    ("overseas_market.json", 10), ("overheating.json", 10), ("sentiment.json", 12),
    ("prediction_log.json", 12), ("sk_hynix_forecast.json", 12), ("japan_crisis.json", 15),
    ("samsung_forecast.json", 12), ("asia_semi.json", 12), ("sk_cycle.json", 14),
    ("earnings.json", 30), ("ai_value_gap.json", 30),
    ("consensus_gap.json", 12), ("semi_decoupling.json", 12), ("yen_sensitivity.json", 12),
    ("event_radar.json", 14),
    # 일일 워크플로
    ("credit_spread.json", 30), ("liquidity_stress.json", 30),
    ("economic_calendar.json", 36), ("m2_data.json", 36),
]


def parse_ts(s):
    """다양한 generated_at 포맷 → aware datetime(KST)."""
    if not s:
        return None
    s = str(s).strip()
    try:
        if "T" in s and ("+" in s or s.endswith("Z")):
            return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(KST)
        s2 = re.sub(r"\s*KST$", "", s)
        return datetime.strptime(s2, "%Y-%m-%d %H:%M:%S").replace(tzinfo=KST)
    except Exception:
        try:
            return datetime.fromisoformat(s).replace(tzinfo=KST)
        except Exception:
            return None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    stale, ok, missing = [], [], []
    for fname, max_h in MONITORED:
        path = os.path.join(DOCS, fname)
        try:
            with open(path, encoding="utf-8") as f:
                d = json.load(f)
        except Exception:
            missing.append(fname)
            continue
        ts = parse_ts(d.get("generated_at") or d.get("asof") or d.get("date"))
        if ts is None:
            continue
        age_h = (now - ts).total_seconds() / 3600
        rec = {"file": fname, "age_h": round(age_h, 1), "max_h": max_h,
               "generated_at": str(d.get("generated_at") or d.get("asof"))}
        if age_h > max_h:
            stale.append(rec)
        else:
            ok.append(rec)

    status = "🔴 정체 감지" if stale else ("🟡 일부 결측" if missing else "🟢 정상")
    print(f"신선도: {status} (정체 {len(stale)} · 정상 {len(ok)} · 결측 {len(missing)})")
    for s in stale:
        print(f"  ⚠️ {s['file']}: {s['age_h']}h 경과 (한계 {s['max_h']}h)")

    # 텔레그램 경보 (정체 시 · 하루 1회 중복방지)
    if stale:
        try:
            from core import send_message, get_secret, load_state, save_state
            today_s = now.strftime("%Y-%m-%d")
            st = load_state("freshness_alert", default={})
            sig = ",".join(sorted(s["file"] for s in stale))
            if get_secret("TELEGRAM_FINANCE_BOT_TOKEN") and st.get("sig") != sig + today_s:
                lines = [f"{s['file']} — {s['age_h']:.0f}h 정체(한계 {s['max_h']}h)" for s in stale[:10]]
                if send_message("🔴 데이터 정체 감지 — 워크플로 갱신 실패 의심\n" + "\n".join(lines)
                                + "\n※ 해당 워크플로 commit/스크립트 점검 필요."):
                    save_state("freshness_alert", {"sig": sig + today_s})
                    print("  🔴 정체 경보 텔레그램 발송")
        except Exception as e:
            print(f"  [WARN] 경보 실패: {e}")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "status": status, "stale_count": len(stale),
        "stale": stale, "ok_count": len(ok), "missing": missing,
        "note": "각 데이터의 generated_at이 예상 주기를 넘으면 정체로 판정·경보. 큐레이션/월간은 제외.",
    }
    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
