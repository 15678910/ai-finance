"""
가치 훼손 모니터 (Value Erosion Monitor)
========================================

브리핑 3.1의 리스크 재정의를 코드로 옮긴 모듈.

  "가치 투자자에게 주가 하락은 주식이 더 싸고 안전해진 것이므로 위험이 아니라
   기회다. 진정한 위험은 기업의 비즈니스 모델이 망가져 내재 가치가 나빠지는 것이다.
   피터 린치: 기업의 가치가 나빠지면 즉시 팔아야 한다."

즉 감시 대상은 **주가가 아니라 펀더멘털**이다. 이 모듈은 주가 하락을 알리지
않는다. 대신 이익·자산·수익성이 실제로 훼손되었는지를 보고, 주가 방향과
교차시켜 네 가지로 분류한다:

  기회 (OPPORTUNITY)     주가↓ · 가치 유지  → 브리핑의 "추가 매수" 구간
  경보 (EROSION)         주가↓ · 가치 훼손  → 싸진 게 아니라 나빠진 것
  은폐 (EROSION_MASKED)  주가↑ · 가치 훼손  → 주가가 훼손을 가리고 있음
  정상 (STABLE)          그 외

비교 기준은 약 1개월 전 스냅샷(docs/value_erosion_state.json)이다. 매일 갱신하면
서서히 진행되는 훼손이 매번 '변화 없음'으로 묻히므로, 기준선은 REBASE_DAYS가
지나야 교체한다.

입력: docs/value_screener.json (value_screener.py가 생성)
출력: docs/value_erosion.json

사용법:
  python value_erosion_monitor.py
  python value_erosion_monitor.py --no-telegram
  python value_erosion_monitor.py --rebase-days 45

🚨 절대 규칙:
  - 시뮬레이션 / 분석 전용
  - 자동 매매 절대 금지
  - 사용자가 직접 검토 후 투자 결정
"""

import os
import sys
import json
import argparse
from datetime import datetime, timezone, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.state_store import load_state, save_state, is_recent_alert, mark_alert_sent
from core.telegram import send_message

KST = timezone(timedelta(hours=9))

SCREENER_FILE = os.path.join(BASE_DIR, "docs", "value_screener.json")
ANALYST_FILE = os.path.join(BASE_DIR, "docs", "analyst_reports.json")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "value_erosion.json")
STATE_NAME = "value_erosion"

# 기준선 교체 주기 — 이 기간이 지나야 비교 스냅샷을 최신값으로 갱신한다.
REBASE_DAYS = 30
# 훼손 판정 임계 (severity 합계)
EROSION_THRESHOLD = 3
# '주가 하락' 판정 임계
PRICE_DROP_PCT = -5.0
# 동일 종목 재알림 방지 (시간)
ALERT_COOLDOWN_HOURS = 24 * 7


# ====================================================================
# 스냅샷 추출
# ====================================================================
def _implied(price, multiple):
    """주가와 배수로부터 주당 지표 역산. PER → EPS, PBR → BPS."""
    if not price or not multiple or multiple <= 0:
        return None
    return price / multiple


def build_snapshot(stock: dict) -> dict:
    """value_screener.json의 종목 1건 → 비교용 스냅샷.

    EPS/BPS를 주가와 배수로 역산해 담는 이유: 스크리너는 PER·PBR만 저장하는데,
    PER은 주가와 이익이 동시에 움직여 그 자체로는 훼손을 판별할 수 없다.
    주가로 나눠 이익(E)과 순자산(B)만 분리해야 "주가가 빠진 것"과
    "이익이 빠진 것"을 구분할 수 있다.
    """
    m = stock.get("metrics", {}) or {}
    price = stock.get("current_price")
    return {
        "name": stock.get("name"),
        "price": price,
        "per": m.get("per"),
        "pbr": m.get("pbr"),
        "per_basis": m.get("per_basis"),
        "eps": _implied(price, m.get("per")),
        "bps": _implied(price, m.get("pbr")),
        "roe_pct": m.get("roe_pct"),
        "operating_margin_pct": m.get("operating_margin_pct"),
        "debt_to_equity": m.get("debt_to_equity"),
        "revenue_growth_pct": m.get("revenue_growth_pct"),
        "trap_flags": [f.get("type") for f in (stock.get("trap_flags") or [])],
        "asof": datetime.now(KST).strftime("%Y-%m-%d"),
    }


def _pct_change(now, before):
    """변화율 (%). 기준값이 0이거나 없으면 None."""
    if now is None or before is None or before == 0:
        return None
    return (now - before) / abs(before) * 100


def _diff(now, before):
    """절대 변화량 (%p 등). 한쪽이라도 없으면 None."""
    if now is None or before is None:
        return None
    return now - before


# ====================================================================
# 훼손 신호 판정
# ====================================================================
def detect_erosion_signals(current: dict, baseline: dict, target_change_pct=None) -> list:
    """펀더멘털 훼손 신호 목록. 각 신호는 severity(1~3)를 가진다.

    severity 3은 그 하나만으로도 경보 임계(3)를 넘긴다 — 이익 자체가 줄어든
    경우가 여기 해당한다. 나머지는 두 개 이상 겹쳐야 경보가 된다.
    """
    signals = []

    # ① 주당 이익 훼손 — 가장 직접적인 가치 훼손
    #    단, PER 기준(후행/선행)이 스냅샷 사이에 바뀌었다면 역산 EPS가
    #    실제 이익 변화가 아닌 지표 교체 때문에 튄 것이므로 판정하지 않는다.
    basis_now = current.get("per_basis")
    basis_before = baseline.get("per_basis")
    basis_stable = not (basis_now and basis_before and basis_now != basis_before)

    eps_chg = _pct_change(current.get("eps"), baseline.get("eps"))
    if basis_stable and eps_chg is not None and eps_chg <= -10:
        signals.append({
            "type": "이익훼손",
            "detail": f"주당순이익 {eps_chg:+.1f}%",
            "severity": 3 if eps_chg <= -20 else 2,
        })

    # ② 순자산 훼손 — 자산가치 기반 저평가의 근거가 무너지는 경우
    bps_chg = _pct_change(current.get("bps"), baseline.get("bps"))
    if bps_chg is not None and bps_chg <= -10:
        signals.append({
            "type": "자산훼손",
            "detail": f"주당순자산 {bps_chg:+.1f}%",
            "severity": 2,
        })

    # ③ 영업이익률 하락 — 비즈니스 모델 경쟁력 약화
    om_diff = _diff(current.get("operating_margin_pct"), baseline.get("operating_margin_pct"))
    if om_diff is not None and om_diff <= -3:
        signals.append({
            "type": "수익성악화",
            "detail": f"영업이익률 {om_diff:+.1f}%p",
            "severity": 2,
        })

    # ④ ROE 하락
    roe_diff = _diff(current.get("roe_pct"), baseline.get("roe_pct"))
    if roe_diff is not None and roe_diff <= -5:
        signals.append({
            "type": "ROE하락",
            "detail": f"ROE {roe_diff:+.1f}%p",
            "severity": 2,
        })

    # ⑤ 부채비율 급증 — 재무 안정성 훼손
    debt_diff = _diff(current.get("debt_to_equity"), baseline.get("debt_to_equity"))
    if debt_diff is not None and debt_diff >= 30:
        signals.append({
            "type": "부채증가",
            "detail": f"부채비율 {debt_diff:+.0f}%p",
            "severity": 1,
        })

    # ⑥ 매출 성장 음전환
    g_now = current.get("revenue_growth_pct")
    g_before = baseline.get("revenue_growth_pct")
    if g_now is not None and g_before is not None and g_before > 0 and g_now < 0:
        signals.append({
            "type": "성장역전",
            "detail": f"매출성장률 {g_before:+.1f}% → {g_now:+.1f}%",
            "severity": 2,
        })

    # ⑦ 가치 함정 플래그 신규 발생 (스크리너의 정성 판정과 연동)
    new_flags = set(current.get("trap_flags") or []) - set(baseline.get("trap_flags") or [])
    for flag in sorted(new_flags):
        signals.append({
            "type": "함정신규",
            "detail": f"가치함정 경고 발생: {flag}",
            "severity": 2,
        })

    # ⑧ 애널리스트 목표주가 컨센서스 하향
    if target_change_pct is not None and target_change_pct <= -10:
        signals.append({
            "type": "컨센서스컷",
            "detail": f"목표주가 컨센서스 {target_change_pct:+.1f}%",
            "severity": 1,
        })

    return signals


def classify(price_change_pct, erosion_score: int) -> tuple:
    """(상태 코드, 한 줄 해석).

    브리핑의 핵심 구분: 주가 하락 그 자체는 신호가 아니다.
    가치가 유지된 채로 빠진 것인지, 가치가 나빠져서 빠진 것인지가 갈림길이다.
    """
    eroded = erosion_score >= EROSION_THRESHOLD
    dropped = price_change_pct is not None and price_change_pct <= PRICE_DROP_PCT

    if eroded and dropped:
        return "EROSION", "가치 훼손 + 주가 하락 — 싸진 것이 아니라 나빠진 것. 매도 검토"
    if eroded and not dropped:
        return "EROSION_MASKED", "가치 훼손 중이나 주가는 견조 — 주가가 훼손을 가리는 중"
    if dropped:
        return "OPPORTUNITY", "가치 유지 + 주가 하락 — 더 싸고 안전해진 구간"
    return "STABLE", "특이 변화 없음"


# ====================================================================
# 데이터 로드
# ====================================================================
def load_screener(path: str = None) -> list:
    """value_screener.json에서 분석 대상 종목 로드 (통과 종목 + 관심 종목)."""
    path = path or SCREENER_FILE
    if not os.path.exists(path):
        print(f"[오류] {os.path.basename(path)} 없음. 먼저 value_screener.py 실행 필요.")
        return []

    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(f"[오류] 스크리너 결과 읽기 실패: {e}")
        return []

    stocks, seen = [], set()
    for group in ("all_passed", "watchlist", "top_picks"):
        for stock in data.get(group) or []:
            ticker = stock.get("ticker")
            if ticker and ticker not in seen:
                seen.add(ticker)
                stocks.append(stock)
    return stocks


def load_analyst_targets(path: str = None) -> dict:
    """종목별 목표주가 컨센서스. {종목코드: 목표주가}"""
    path = path or ANALYST_FILE
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}

    targets = {}
    for stock in data.get("stocks", []):
        code, target = stock.get("code"), stock.get("consensus_target")
        if code and target:
            targets[code] = target
    return targets


def _baseline_age_days(baseline: dict, today: datetime = None) -> int:
    """기준선 스냅샷 경과일. 파싱 실패 시 0 (교체하지 않음)."""
    asof = (baseline or {}).get("asof")
    if not asof:
        return 0
    try:
        asof_date = datetime.strptime(asof, "%Y-%m-%d").replace(tzinfo=KST)
        return ((today or datetime.now(KST)) - asof_date).days
    except ValueError:
        return 0


# ====================================================================
# 분석
# ====================================================================
def analyze(stocks: list, state: dict, targets: dict,
            rebase_days: int = REBASE_DAYS, today: datetime = None) -> tuple:
    """전 종목 훼손 판정. (결과 리스트, 갱신된 baselines)"""
    baselines = dict(state.get("baselines", {}))
    results = []

    for stock in stocks:
        ticker = stock.get("ticker")
        if not ticker:
            continue

        current = build_snapshot(stock)
        baseline = baselines.get(ticker)

        # 최초 관측 — 기준선만 세우고 판정은 다음 실행부터
        if not baseline:
            current["consensus_target"] = targets.get(ticker)
            baselines[ticker] = current
            results.append({
                "ticker": ticker,
                "name": current["name"],
                "status": "NEW",
                "interpretation": "기준선 최초 기록 — 다음 실행부터 비교",
                "erosion_score": 0,
                "signals": [],
                "price_change_pct": None,
                "baseline_asof": current["asof"],
            })
            continue

        # 목표주가 컨센서스 변화 (기준선에 저장된 값과 비교)
        target_change = _pct_change(targets.get(ticker), baseline.get("consensus_target"))

        signals = detect_erosion_signals(current, baseline, target_change)
        erosion_score = sum(s["severity"] for s in signals)
        price_change = _pct_change(current.get("price"), baseline.get("price"))
        status, interpretation = classify(price_change, erosion_score)

        results.append({
            "ticker": ticker,
            "name": current["name"],
            "status": status,
            "interpretation": interpretation,
            "erosion_score": erosion_score,
            "signals": signals,
            "price_change_pct": round(price_change, 2) if price_change is not None else None,
            "eps_change_pct": round(_pct_change(current.get("eps"), baseline.get("eps")), 2)
                              if _pct_change(current.get("eps"), baseline.get("eps")) is not None else None,
            "baseline_asof": baseline.get("asof"),
            "current": current,
        })

        # 기준선 교체 — REBASE_DAYS 경과 시에만
        if _baseline_age_days(baseline, today) >= rebase_days:
            current["consensus_target"] = targets.get(ticker)
            baselines[ticker] = current

    return results, baselines


def build_message(alerts: list, all_results: list = None) -> str:
    """텔레그램 알림 본문. 경보 종목이 없으면 빈 문자열.

    alerts는 쿨다운을 통과한 경보 종목, all_results는 참고용 전체 결과
    (기회 구간 요약에 사용).
    """
    if not alerts:
        return ""

    alerts = sorted(alerts, key=lambda r: r["erosion_score"], reverse=True)
    lines = ["🔻 가치 훼손 경보", "=" * 25, "",
             "주가가 아니라 펀더멘털이 나빠진 종목입니다.", ""]

    for r in alerts[:10]:
        icon = "🚨" if r["status"] == "EROSION" else "⚠️"
        lines.append(f"{icon} {r['name']} ({r['ticker']}) — 훼손도 {r['erosion_score']}")
        for s in r["signals"]:
            lines.append(f"   · {s['type']}: {s['detail']}")
        if r["price_change_pct"] is not None:
            lines.append(f"   주가 {r['price_change_pct']:+.1f}% (기준 {r['baseline_asof']})")
        lines.append(f"   → {r['interpretation']}")
        lines.append("")

    opportunities = [r for r in (all_results or alerts) if r["status"] == "OPPORTUNITY"]
    if opportunities:
        names = ", ".join(f"{r['name']}({r['price_change_pct']:+.0f}%)" for r in opportunities[:5])
        lines.append(f"📉 가치 유지 · 주가 하락: {names}")
        lines.append("")

    lines.append("🚨 시뮬레이션/분석용. 자동 매매 금지.")
    lines.append("대시보드: https://15678910.github.io/ai-finance/")
    return "\n".join(lines)


# ====================================================================
# 메인
# ====================================================================
def main():
    parser = argparse.ArgumentParser(description="가치 훼손 모니터")
    parser.add_argument("--no-telegram", action="store_true", help="텔레그램 전송 생략")
    parser.add_argument("--force", action="store_true", help="쿨다운 무시하고 알림")
    parser.add_argument("--rebase-days", type=int, default=REBASE_DAYS,
                        help=f"기준선 교체 주기 (기본 {REBASE_DAYS}일)")
    args = parser.parse_args()

    now = datetime.now(KST)
    print("=" * 65)
    print("  가치 훼손 모니터 (Value Erosion Monitor)")
    print(f"  시각: {now.strftime('%Y-%m-%d %H:%M:%S')} KST")
    print(f"  기준선 교체 주기: {args.rebase_days}일")
    print("=" * 65)

    stocks = load_screener()
    if not stocks:
        return 1
    print(f"\n  대상 종목: {len(stocks)}개")

    state = load_state(STATE_NAME, {"baselines": {}})
    targets = load_analyst_targets()
    if targets:
        print(f"  목표주가 컨센서스: {len(targets)}개 종목")

    results, baselines = analyze(stocks, state, targets, args.rebase_days, now)

    counts = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    print(f"\n[판정 결과]")
    for status, label in [("EROSION", "🚨 가치 훼손"), ("EROSION_MASKED", "⚠️ 훼손 은폐"),
                          ("OPPORTUNITY", "📉 기회 (가치 유지·주가 하락)"),
                          ("STABLE", "· 정상"), ("NEW", "· 신규 기준선")]:
        if counts.get(status):
            print(f"  {label}: {counts[status]}개")

    for r in results:
        if r["status"] in ("EROSION", "EROSION_MASKED"):
            print(f"\n  {r['name']} ({r['ticker']}) — 훼손도 {r['erosion_score']}")
            for s in r["signals"]:
                print(f"    · {s['type']}: {s['detail']}")
            print(f"    → {r['interpretation']}")

    # 저장
    output = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "rebase_days": args.rebase_days,
        "erosion_threshold": EROSION_THRESHOLD,
        "counts": counts,
        "results": sorted(results, key=lambda r: r["erosion_score"], reverse=True),
        "note": ("주가가 아닌 펀더멘털(EPS·BPS·수익성·재무)의 훼손을 감시한다. "
                 "주가 하락은 가치가 유지되는 한 위험이 아니라 기회로 분류한다."),
        "warning": "🚨 시뮬레이션/분석용. 자동 매매 절대 금지. 사용자 직접 검토 후 투자 결정.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    state["baselines"] = baselines
    state["last_run"] = now.isoformat()

    # 텔레그램 (종목별 쿨다운)
    if not args.no_telegram:
        alerts = [r for r in results if r["status"] in ("EROSION", "EROSION_MASKED")]
        fresh = [r for r in alerts
                 if args.force or not is_recent_alert(state, f"erosion:{r['ticker']}",
                                                      ALERT_COOLDOWN_HOURS)]
        if fresh:
            message = build_message(fresh, results)
            if message and send_message(message):
                for r in fresh:
                    state = mark_alert_sent(state, f"erosion:{r['ticker']}")
                print(f"  [텔레그램] 전송 완료 ({len(fresh)}개 종목)")
            else:
                print("  [텔레그램] 설정 없음 또는 전송 실패")
        elif alerts:
            print(f"  [텔레그램] 쿨다운 중 ({len(alerts)}개 종목) — 전송 생략")
        else:
            print("  [텔레그램] 경보 없음 — 전송 생략")

    save_state(STATE_NAME, state)

    print("\n" + "=" * 65)
    print("  ⚠️ 본 결과는 시뮬레이션 전용입니다.")
    print("  ⚠️ 실제 투자 결정은 본인의 판단 필요.")
    print("=" * 65)
    return 0


if __name__ == "__main__":
    sys.exit(main())
