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
    ("earnings.json", 30), ("ai_value_gap.json", 30), ("fx_swap.json", 14),
    ("fx_intervention.json", 40),   # REER 월간·FIMA 주간이라 여유 있게
    ("money_macro.json", 40),       # OECD M3 월간·BIS 주택 분기
    ("analyst_reports.json", 30), ("watchlist.json", 14), ("dart_disclosures.json", 16),
    ("vkospi.json", 14), ("key_events.json", 30), ("composite_signal.json", 12),
    ("prediction_quality.json", 14), ("memory_price.json", 14), ("quotes.json", 14),
    ("consensus_gap.json", 12), ("semi_decoupling.json", 12), ("yen_sensitivity.json", 12),
    ("event_radar.json", 14),
    # 일일 워크플로
    ("credit_spread.json", 30), ("liquidity_stress.json", 30),
    ("economic_calendar.json", 36), ("m2_data.json", 36), ("debt_gdp.json", 40),
    ("fear_greed.json", 30), ("trend_signal.json", 36), ("guru_13f.json", 36),
    ("kospi_open_log.json", 30),   # 시초가 예측 워크포워드 채점 (해외시장 워크플로에 동봉, 매시간)
    ("dart_financials.json", 40),
    # 2026-11-03 추가: 아침 브리핑·투자자·게이트·시가총액·기계경제 — 그동안 감시 밖이라 3일 정체(8-class)와 '빈 데이터'(투자자 매매)를 놓쳤다.
    ("morning_report.json", 27), ("investor_flow.json", 30), ("market_investor_breakdown.json", 30),
    ("market_breadth.json", 30), ("flow_streak.json", 30),
    ("market_cap_tracker.json", 8), ("machine_economy.json", 8),
    # kospi_valuation.json: 원천(INDEXerGO)이 GitHub 데이터센터 IP를 지역차단(403) → 자동갱신 불가.
    #   PER/PBR은 완만한 지표라 저빈도 허용 + 거짓 경보 방지 위해 감시 제외. (필요 시 수동 갱신)
]

# 평일에만 생성되는 파일 — 주말·월요일 오전에는 금요일 자가 최신이므로 한계를 늘려 거짓 경보를 막는다
WEEKDAY_ONLY = {"investor_flow.json", "market_investor_breakdown.json", "market_breadth.json", "flow_streak.json", "morning_report.json"}


def weekend_allowance_h(now):
    """토·일·월 20:00 전이면 금요일 저녁 이후 지난 시간을 한계에 더한다."""
    wd = now.weekday()
    if wd == 5:
        return 24 + now.hour
    if wd == 6:
        return 48 + now.hour
    if wd == 0 and now.hour < 20:
        return 72
    return 0


# '생성 시각은 최신인데 내용이 비어 있는' 실패 탐지 — 2026-10-30~11-02 투자자 매매(분석 0종목)가 이 유형이었다.
def _chk_investor_flow(d, now):
    n = d.get("analyzed_count") or 0
    return (n > 0, f"분석 {n}종목")


def _chk_breakdown(d, now):
    if d.get("stale"):
        return (False, "직전 데이터 유지(stale)")
    m = d.get("markets") or {}
    n = sum(len((v or {}).get("rows") or []) for v in m.values())
    return (n > 0, f"행 {n}")


def _chk_breadth(d, now):
    n = ((d.get("breadth") or {}).get("n")) or 0
    return (n >= 100, f"종목 {n}")


def _chk_flow_streak(d, now):
    n = ((d.get("counts") or {}).get("fetched")) or 0
    return (n >= 100, f"수집 {n}종목")


def _chk_morning(d, now):
    # 09:00 KST 이후에는 오늘 자여야 한다 (예약 지연·미실행 탐지)
    if now.hour >= 9 and d.get("date") != now.strftime("%Y-%m-%d") and now.weekday() < 5:
        return (False, f"오늘 자 없음 (최신 {d.get('date')})")
    return (True, f"기준일 {d.get('date')}")


def _chk_market_cap(d, now):
    cards = d.get("cards") or []
    bad = [c["key"] for c in cards if c.get("quality") in ("unavailable", "stale")]
    return (len(bad) < len(cards) if cards else False, f"카드 {len(cards)-len(bad)}/{len(cards)} 정상" + (f" (문제: {','.join(bad)})" if bad else ""))


def _chk_machine_economy(d, now):
    st = ((d.get("facts") or {}).get("stale")) or []
    return (len(st) < 2, f"stale {st}" if st else "정상")


CONTENT_CHECKS = {
    "investor_flow.json": _chk_investor_flow, "market_investor_breakdown.json": _chk_breakdown,
    "market_breadth.json": _chk_breadth, "flow_streak.json": _chk_flow_streak,
    "morning_report.json": _chk_morning, "market_cap_tracker.json": _chk_market_cap,
    "machine_economy.json": _chk_machine_economy,
}


def evaluate(files, now, loader):
    """[(fname, max_h)] → (stale, ok, missing, empty). loader(fname) → dict | None. 순수 함수(테스트용)."""
    stale, ok, missing, empty = [], [], [], []
    allowance = weekend_allowance_h(now)
    for fname, max_h in files:
        d = loader(fname)
        if d is None:
            missing.append(fname)
            continue
        ts = parse_ts(d.get("generated_at") or d.get("asof") or d.get("date"))
        if ts is None:
            continue
        age_h = (now - ts).total_seconds() / 3600
        limit = max_h + (allowance if fname in WEEKDAY_ONLY else 0)
        rec = {"file": fname, "age_h": round(age_h, 1), "max_h": limit,
               "generated_at": str(d.get("generated_at") or d.get("asof"))}
        chk = CONTENT_CHECKS.get(fname)
        if chk:
            try:
                good, why = chk(d, now)
            except Exception as e:
                good, why = True, f"검사 오류 {type(e).__name__}"
            rec["content"] = why
            if not good:
                rec["reason"] = f"내용 이상: {why}"
                empty.append(rec)
                stale.append(rec)          # 배너·경보는 정체와 같은 경로로
                continue
        if age_h > limit:
            rec["reason"] = f"{rec['age_h']}h 경과 (한계 {limit}h)"
            stale.append(rec)
        else:
            ok.append(rec)
    return stale, ok, missing, empty


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

    def loader(fname):
        try:
            with open(os.path.join(DOCS, fname), encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    stale, ok, missing, empty = evaluate(MONITORED, now, loader)

    status = "🔴 정체 감지" if stale else ("🟡 일부 결측" if missing else "🟢 정상")
    print(f"신선도: {status} (정체/이상 {len(stale)} · 정상 {len(ok)} · 결측 {len(missing)} · 내용 이상 {len(empty)})")
    for s in stale:
        print(f"  ⚠️ {s['file']}: {s.get('reason')}")

    # 텔레그램 경보 (정체 시 · 하루 1회 중복방지)
    if stale:
        try:
            from core import send_message, get_secret, load_state, save_state
            today_s = now.strftime("%Y-%m-%d")
            st = load_state("freshness_alert", default={})
            sig = ",".join(sorted(s["file"] for s in stale))
            if get_secret("TELEGRAM_FINANCE_BOT_TOKEN") and st.get("sig") != sig + today_s:
                lines = [f"{s['file']} — {s.get('reason')}" for s in stale[:12]]
                if send_message("🔴 데이터 정체·이상 감지 — 워크플로 점검 필요\n" + "\n".join(lines)
                                + "\n※ 해당 워크플로 실행 기록(Actions)과 스크립트 로그를 확인하세요."):
                    save_state("freshness_alert", {"sig": sig + today_s})
                    print("  🔴 정체 경보 텔레그램 발송")
        except Exception as e:
            print(f"  [WARN] 경보 실패: {e}")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "status": status, "stale_count": len(stale),
        "stale": stale, "ok_count": len(ok), "missing": missing, "empty": empty,
        "note": "각 데이터의 generated_at이 예상 주기를 넘거나(정체) 내용이 비어 있으면(이상) 판정·경보. 평일 전용 파일은 주말 유예. 큐레이션/월간은 제외.",
    }
    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
