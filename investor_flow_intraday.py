"""
한국 시장 장중 투자자 매매 모니터 (Phase B — 시간별 모멘텀)
=============================================================

목적:
  네이버 금융 종목별 외국인·기관 매매 페이지는 장중에도 누적 데이터를 반환.
  매시간 스냅샷을 저장하고 시간별 변화(모멘텀)를 추적하여 장중 강한 매수/매도
  가속을 즉시 감지.

Phase A와의 차이:
  · Phase A: 매일 16:30 일별 데이터 (장 마감 후)
  · Phase B: 장중 매시간 09:30, 10:30 ... 15:30 (KST) 시간별 누적
  · Phase B 비교 단위: 현재 vs 1시간 전 = 1시간 매매 가속도

자동 감지 시그널:
  ⭐⭐⭐⭐⭐ 외인 신규 강매수 가속 (이전 시간 대비 +50k주 이상 추가 매수)
  ⭐⭐⭐⭐⭐ 외인 매수→매도 전환 (이전 시간엔 +, 이번엔 -)
  ⭐⭐⭐⭐  기관 강매수 가속
  ⭐⭐⭐   외인 매도 강도 가속 (-50k주 이상 추가 매도)

스냅샷 저장 (ring buffer): 최근 8개 (≈하루 분량) 보관

🚨 시뮬레이션. 자동 매매 금지.
"""

import os
import json
import time
from datetime import datetime, timezone, timedelta

from core import send_message, load_state, save_state

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "investor_flow_intraday.json")
STATE_NAME = "investor_flow_intraday"
KST = timezone(timedelta(hours=9))

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
TIMEOUT_SEC = 15
SNAPSHOT_RETENTION = 8  # 시간별 ring buffer

# 추적 대상 — Phase A와 동일 universe (재import 회피용으로 명시)
from investor_flow_monitor import TARGET_STOCKS, fetch_naver_flow  # noqa


# ====================================================================
# 시간별 모멘텀 분석
# ====================================================================
def compute_momentum(current: dict, previous: dict | None) -> dict:
    """현재 스냅샷 vs 1시간 전 스냅샷 → 모멘텀 시그널."""
    if not previous:
        return {"signals": [], "delta_foreign": None, "delta_institutional": None, "first_snapshot": True}

    cur_f = current.get("foreign_today_net_shares", 0) or 0
    cur_i = current.get("institutional_today_net_shares", 0) or 0
    prev_f = previous.get("foreign_today_net_shares", 0) or 0
    prev_i = previous.get("institutional_today_net_shares", 0) or 0

    delta_f = cur_f - prev_f  # 이번 1시간 동안 외인 추가 순매수
    delta_i = cur_i - prev_i

    signals = []

    # 1. 외인 신규 강매수 가속 (1시간 +50k주 이상)
    if delta_f >= 50000:
        signals.append({
            "type": "foreign_buy_accel", "strength": 5, "emoji": "🚀",
            "label": "외인 매수 가속",
            "description": f"1시간 +{delta_f:,}주 (누적 {cur_f:+,})",
        })
    # 2. 외인 매수→매도 전환
    elif prev_f > 0 and cur_f < 0:
        signals.append({
            "type": "foreign_flip_sell", "strength": 5, "emoji": "🔻",
            "label": "외인 매수→매도 전환",
            "description": f"이전 {prev_f:+,} → 현재 {cur_f:+,}",
        })
    # 3. 외인 매도→매수 전환
    elif prev_f < 0 and cur_f > 0:
        signals.append({
            "type": "foreign_flip_buy", "strength": 5, "emoji": "🔺",
            "label": "외인 매도→매수 전환",
            "description": f"이전 {prev_f:+,} → 현재 {cur_f:+,}",
        })
    # 4. 외인 매도 강도 가속
    elif delta_f <= -50000:
        signals.append({
            "type": "foreign_sell_accel", "strength": 4, "emoji": "🩸",
            "label": "외인 매도 가속",
            "description": f"1시간 {delta_f:,}주 (누적 {cur_f:+,})",
        })

    # 기관 동향
    if delta_i >= 50000:
        signals.append({
            "type": "institutional_buy_accel", "strength": 4, "emoji": "🔥",
            "label": "기관 매수 가속",
            "description": f"1시간 +{delta_i:,}주",
        })
    elif delta_i <= -50000:
        signals.append({
            "type": "institutional_sell_accel", "strength": 3, "emoji": "📉",
            "label": "기관 매도 가속",
            "description": f"1시간 {delta_i:,}주",
        })

    max_strength = max((s["strength"] for s in signals), default=0)
    return {
        "signals": signals,
        "delta_foreign": delta_f,
        "delta_institutional": delta_i,
        "max_strength": max_strength,
    }


def take_snapshot() -> dict:
    """현재 시점 스냅샷 — 모든 종목의 today 행 (장중 누적) 추출."""
    snapshot_time = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    print(f"\n[{snapshot_time} KST 스냅샷]")
    stocks = []
    for s in TARGET_STOCKS:
        ticker = s["ticker"]
        rows = fetch_naver_flow(ticker)
        if not rows:
            continue
        # 가장 최신 행 = today (장중) 또는 yesterday (장 외)
        today_row = sorted(rows, key=lambda r: r["date"], reverse=True)[0]
        stocks.append({
            "ticker": ticker,
            "name": s["name"],
            "sector": s.get("sector"),
            "today_date": today_row["date"],
            "today_close": today_row["close"],
            "today_change_pct": today_row["change_pct"],
            "foreign_today_net_shares": today_row["foreign_net_shares"],
            "institutional_today_net_shares": today_row["institutional_net_shares"],
            "foreign_holding_pct": today_row["foreign_holding_pct"],
        })
        time.sleep(0.2)
    print(f"  · {len(stocks)}개 종목 수집")
    return {
        "snapshot_time": snapshot_time,
        "stocks": stocks,
    }


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 72)
    print("  장중 투자자 매매 모니터 (Phase B — 시간별 모멘텀)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 72)

    # State: 직전 스냅샷 + 알림 이력
    state = load_state(STATE_NAME, default={"snapshots": [], "last_alert_keys": []})
    snapshots = state.get("snapshots", [])

    # 새 스냅샷 수집
    new_snapshot = take_snapshot()

    # 직전 스냅샷과 비교 (모멘텀 계산)
    previous_snapshot = snapshots[-1] if snapshots else None

    momentum_results = []
    if previous_snapshot:
        # 종목 매칭 (ticker로 join)
        prev_map = {s["ticker"]: s for s in previous_snapshot.get("stocks", [])}
        for cur_stock in new_snapshot["stocks"]:
            prev_stock = prev_map.get(cur_stock["ticker"])
            mom = compute_momentum(cur_stock, prev_stock)
            momentum_results.append({**cur_stock, "momentum": mom})
    else:
        # 첫 스냅샷 — 모멘텀 없음
        for cur_stock in new_snapshot["stocks"]:
            momentum_results.append({**cur_stock, "momentum": {"signals": [], "first_snapshot": True}})

    # 시그널 강도 내림차순 정렬
    momentum_results.sort(key=lambda r: r.get("momentum", {}).get("max_strength", 0), reverse=True)

    # 시그널별 Top
    strong_signals = [r for r in momentum_results if r.get("momentum", {}).get("max_strength", 0) >= 4]
    foreign_buy_accel = [r for r in momentum_results if any(s["type"] == "foreign_buy_accel"
                                                             for s in r.get("momentum", {}).get("signals", []))]
    foreign_sell_accel = [r for r in momentum_results if any(s["type"] == "foreign_sell_accel"
                                                              for s in r.get("momentum", {}).get("signals", []))]
    foreign_flips = [r for r in momentum_results if any(s["type"] in ("foreign_flip_buy", "foreign_flip_sell")
                                                         for s in r.get("momentum", {}).get("signals", []))]
    inst_buy_accel = [r for r in momentum_results if any(s["type"] == "institutional_buy_accel"
                                                          for s in r.get("momentum", {}).get("signals", []))]

    # 신규 알림 (이전 스냅샷에서 알린 종목은 제외)
    last_alert_keys = set(state.get("last_alert_keys", []))
    new_alerts = []
    current_alert_keys = set()
    for r in strong_signals:
        key = f"{r['ticker']}:{':'.join(s['type'] for s in r['momentum']['signals'])}"
        current_alert_keys.add(key)
        if key not in last_alert_keys:
            new_alerts.append(r)

    if new_alerts:
        lines = ["⚡ 장중 매매 모멘텀 — 강력 시그널", "=" * 30, ""]
        lines.append(f"📅 {new_snapshot['snapshot_time']} KST 기준")
        lines.append("")
        for r in new_alerts[:10]:
            sigs = r["momentum"]["signals"]
            primary = sigs[0]
            lines.append(f"{primary['emoji']} {r['name']} ({r['ticker']}) — {r['sector']}")
            for s in sigs[:2]:
                lines.append(f"   {s['label']}: {s['description']}")
            lines.append(f"   종가 {r['today_close']:,}원 ({r['today_change_pct']:+.2f}%) · 외인보유 {r['foreign_holding_pct']:.2f}%")
            lines.append("")
        lines.append("🚨 시뮬레이션. 자동 매매 금지.")
        lines.append("대시보드: https://15678910.github.io/ai-finance/")
        try:
            send_message("\n".join(lines))
            print(f"\n  ✅ 텔레그램 발송: 신규 강력 시그널 {len(new_alerts)}건")
        except Exception as e:
            print(f"\n  ❌ 텔레그램 발송 실패: {e}")

    # State 업데이트 — ring buffer
    snapshots.append(new_snapshot)
    if len(snapshots) > SNAPSHOT_RETENTION:
        snapshots = snapshots[-SNAPSHOT_RETENTION:]
    state["snapshots"] = snapshots
    state["last_alert_keys"] = list(current_alert_keys)
    save_state(STATE_NAME, state)

    # 시간별 누적 차트 데이터 (외인 합계 시계열)
    timeline = []
    for snap in snapshots:
        total_foreign = sum(s.get("foreign_today_net_shares", 0) or 0 for s in snap.get("stocks", []))
        total_inst = sum(s.get("institutional_today_net_shares", 0) or 0 for s in snap.get("stocks", []))
        timeline.append({
            "time": snap.get("snapshot_time"),
            "total_foreign_net_shares": total_foreign,
            "total_institutional_net_shares": total_inst,
        })

    # 결과 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "current_snapshot_time": new_snapshot["snapshot_time"],
        "previous_snapshot_time": previous_snapshot.get("snapshot_time") if previous_snapshot else None,
        "methodology": (
            "장중 매시간 네이버 금융 스냅샷 → 종목별 1시간 매매 가속도 + 매수/매도 전환 감지. "
            "ring buffer 최근 8개 스냅샷 보관."
        ),
        "snapshot_count": len(snapshots),
        "analyzed_count": len(momentum_results),
        "strong_signals_count": len(strong_signals),
        "new_alerts_this_run": len(new_alerts),
        "is_first_snapshot": previous_snapshot is None,
        "results": momentum_results[:30],  # 강도 내림차순 Top 30
        "top_signals": {
            "foreign_buy_accel": [_summary(r) for r in foreign_buy_accel[:10]],
            "foreign_sell_accel": [_summary(r) for r in foreign_sell_accel[:10]],
            "foreign_flips": [_summary(r) for r in foreign_flips[:10]],
            "institutional_buy_accel": [_summary(r) for r in inst_buy_accel[:10]],
        },
        "market_timeline": timeline,  # 시간별 시장 전체 외인·기관 누적
        "data_source": "네이버 금융 (시간별 스냅샷)",
        "warning": "🚨 장중 데이터. 갱신 지연 가능. 시뮬레이션 전용.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    print("\n" + "=" * 72)
    print(f"  완료: 스냅샷 {len(snapshots)}개 보관 · 강력 시그널 {len(strong_signals)}건 · 신규 알림 {len(new_alerts)}건")
    print("=" * 72)


def _summary(r: dict) -> dict:
    mom = r.get("momentum", {})
    sigs = mom.get("signals", [])
    return {
        "ticker": r["ticker"],
        "name": r["name"],
        "sector": r.get("sector"),
        "today_close": r.get("today_close"),
        "today_change_pct": r.get("today_change_pct"),
        "foreign_today_net_shares": r.get("foreign_today_net_shares"),
        "institutional_today_net_shares": r.get("institutional_today_net_shares"),
        "delta_foreign_1h": mom.get("delta_foreign"),
        "delta_institutional_1h": mom.get("delta_institutional"),
        "primary_signal": sigs[0]["label"] if sigs else None,
        "primary_signal_emoji": sigs[0]["emoji"] if sigs else None,
        "foreign_holding_pct": r.get("foreign_holding_pct"),
    }


if __name__ == "__main__":
    main()
