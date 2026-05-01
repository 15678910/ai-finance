"""
Forward Test 평가 시스템 (실제 성과 검증)
==========================================

과거 추천한 Top 10 종목들이 실제로 KOSPI를 이겼는지 검증합니다.
1주, 1개월, 3개월, 6개월 시점의 실제 수익률을 추적합니다.

황금 지표:
  - 실현 알파 (Realized Alpha): 추천 종목 평균 수익률 - KOSPI 수익률
  - 실현 Hit Rate: KOSPI 이긴 종목 비율
  - 누적 알파: 모든 추천의 가중 평균 알파

🚨 시뮬레이션/분석용. 자동 매매 절대 금지.
"""

import os
import sys
import json
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta

try:
    import yfinance as yf
    import pandas as pd
    import numpy as np
except ImportError as e:
    print(f"[오류] 라이브러리 미설치: {e}")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
HISTORY_FILE = os.path.join(BASE_DIR, "docs", "forward_test_history.json")
EVAL_FILE = os.path.join(BASE_DIR, "docs", "forward_test_evaluation.json")

KOSPI_TICKER = "^KS11"
KST = timezone(timedelta(hours=9))

# 평가 시점 (일)
EVAL_HORIZONS = [7, 30, 90, 180]


# ====================================================================
# 수익률 수집
# ====================================================================
def get_return_between(ticker: str, start_date: str, end_date: str = None) -> float:
    """특정 기간의 수익률 (%)."""
    if end_date is None:
        end_date = datetime.now().strftime("%Y-%m-%d")

    try:
        # KR 종목 자동 처리
        if ticker.isdigit() and len(ticker) == 6:
            for suffix in [".KS", ".KQ"]:
                yf_ticker = ticker + suffix
                stock = yf.Ticker(yf_ticker)
                hist = stock.history(start=start_date, end=end_date)
                if not hist.empty and len(hist) >= 2:
                    break
        else:
            stock = yf.Ticker(ticker)
            hist = stock.history(start=start_date, end=end_date)

        if hist.empty or len(hist) < 2:
            return None

        start = hist["Close"].iloc[0]
        end = hist["Close"].iloc[-1]
        if start <= 0:
            return None

        return round((end - start) / start * 100, 2)
    except Exception:
        return None


def get_kospi_return(start_date: str, end_date: str = None) -> float:
    """KOSPI 수익률."""
    return get_return_between(KOSPI_TICKER, start_date, end_date) or 0


# ====================================================================
# 평가
# ====================================================================
def evaluate_record(record: dict) -> dict:
    """단일 추천 기록의 실제 성과 평가."""
    rec_date_str = record["date"].split()[0]  # "2026-05-01"
    rec_date = datetime.strptime(rec_date_str, "%Y-%m-%d")
    today = datetime.now()
    days_elapsed = (today - rec_date).days

    if days_elapsed < EVAL_HORIZONS[0]:
        return None  # 너무 빠름

    # 평가할 horizon 선택 (가장 가까운 기준일)
    horizons_to_eval = [h for h in EVAL_HORIZONS if h <= days_elapsed]
    if not horizons_to_eval:
        return None

    evaluations = {}
    top_picks = record.get("top_picks", [])

    for horizon in horizons_to_eval:
        horizon_end = (rec_date + timedelta(days=horizon)).strftime("%Y-%m-%d")
        # KOSPI 수익률
        kospi_ret = get_kospi_return(rec_date_str, horizon_end)

        # 각 추천 종목 수익률
        stock_returns = []
        for pick in top_picks:
            ticker = pick["ticker"]
            ret = get_return_between(ticker, rec_date_str, horizon_end)
            if ret is not None:
                stock_returns.append({
                    "ticker": ticker,
                    "name": pick["name"],
                    "return_pct": ret,
                    "alpha_pct": round(ret - kospi_ret, 2),
                    "won": ret > kospi_ret,
                })

        if not stock_returns:
            continue

        avg_return = round(np.mean([s["return_pct"] for s in stock_returns]), 2)
        wins = sum(1 for s in stock_returns if s["won"])
        hit_rate = round(wins / len(stock_returns) * 100, 1)
        realized_alpha = round(avg_return - kospi_ret, 2)

        evaluations[f"{horizon}d"] = {
            "horizon_days": horizon,
            "evaluation_date": horizon_end,
            "avg_return_pct": avg_return,
            "kospi_return_pct": kospi_ret,
            "realized_alpha_pct": realized_alpha,
            "hit_rate_pct": hit_rate,
            "winners": wins,
            "total": len(stock_returns),
            "stock_details": stock_returns,
        }

    return {
        "recommendation_date": record["date"],
        "expected_alpha_pct": record.get("expected_alpha"),
        "expected_hit_rate_pct": record.get("expected_hit_rate"),
        "weights": record.get("weights"),
        "days_elapsed": days_elapsed,
        "horizons": evaluations,
    }


# ====================================================================
# 누적 통계
# ====================================================================
def calculate_cumulative_stats(evaluations: list) -> dict:
    """전체 평가의 누적 통계."""
    if not evaluations:
        return {}

    stats = {}
    for horizon in EVAL_HORIZONS:
        key = f"{horizon}d"
        # 해당 horizon에 평가가 있는 기록만
        relevant = [e["horizons"].get(key) for e in evaluations if e and key in e.get("horizons", {})]
        relevant = [r for r in relevant if r]

        if not relevant:
            continue

        avg_alpha = round(np.mean([r["realized_alpha_pct"] for r in relevant]), 2)
        std_alpha = round(np.std([r["realized_alpha_pct"] for r in relevant]), 2)
        avg_hit_rate = round(np.mean([r["hit_rate_pct"] for r in relevant]), 1)
        positive_alpha_count = sum(1 for r in relevant if r["realized_alpha_pct"] > 0)
        win_rate = round(positive_alpha_count / len(relevant) * 100, 1)

        # Information Ratio = avg_alpha / std_alpha
        info_ratio = round(avg_alpha / std_alpha, 2) if std_alpha > 0 else 0

        stats[key] = {
            "horizon_days": horizon,
            "evaluations_count": len(relevant),
            "avg_realized_alpha_pct": avg_alpha,
            "std_alpha_pct": std_alpha,
            "avg_hit_rate_pct": avg_hit_rate,
            "win_rate_pct": win_rate,  # 알파 > 0 비율
            "information_ratio": info_ratio,
        }

    return stats


def calculate_learning_curve(evaluations: list) -> list:
    """진화 학습 곡선: 시간에 따른 알파 추세."""
    timeline = []
    for ev in sorted(evaluations, key=lambda x: x["recommendation_date"]):
        if not ev or "30d" not in ev.get("horizons", {}):
            continue
        h30 = ev["horizons"]["30d"]
        timeline.append({
            "date": ev["recommendation_date"],
            "expected_alpha": ev.get("expected_alpha_pct"),
            "realized_alpha_30d": h30["realized_alpha_pct"],
            "hit_rate_30d": h30["hit_rate_pct"],
        })
    return timeline


# ====================================================================
# 텔레그램
# ====================================================================
def send_telegram_summary(stats: dict, n_evals: int):
    env_path = os.path.join(CONFIG_DIR, ".env")
    bot_token, chat_id = None, None
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                if "=" not in line or line.startswith("#"):
                    continue
                k, v = line.strip().split("=", 1)
                v = v.strip().strip("'\"")
                if k.strip() == "TELEGRAM_FINANCE_BOT_TOKEN":
                    bot_token = v
                elif k.strip() == "TELEGRAM_FINANCE_CHAT_ID":
                    chat_id = v
    bot_token = bot_token or os.environ.get("TELEGRAM_FINANCE_BOT_TOKEN")
    chat_id = chat_id or os.environ.get("TELEGRAM_FINANCE_CHAT_ID")

    if not bot_token or not chat_id or not stats:
        return

    lines = [
        "📊 Forward Test 평가",
        "=" * 25,
        "",
        f"평가 추천 수: {n_evals}건",
        "",
    ]

    for horizon_key, s in stats.items():
        days = s["horizon_days"]
        lines.append(f"[{days}일 시점]")
        lines.append(f"  실현 알파: {s['avg_realized_alpha_pct']:+.2f}% (표준편차 {s['std_alpha_pct']:.2f})")
        lines.append(f"  평균 Hit Rate: {s['avg_hit_rate_pct']}%")
        lines.append(f"  KOSPI 이긴 추천: {s['win_rate_pct']}% ({int(s['win_rate_pct']/100*s['evaluations_count'])}/{s['evaluations_count']}건)")
        lines.append(f"  Information Ratio: {s['information_ratio']}")
        lines.append("")

    lines.append("🚨 시뮬레이션. 자동 매매 금지.")
    lines.append("\n대시보드: https://15678910.github.io/ai-finance/")

    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        body = urllib.parse.urlencode({"chat_id": chat_id, "text": "\n".join(lines)}).encode()
        req = urllib.request.Request(url, data=body, method="POST")
        with urllib.request.urlopen(req, timeout=10) as resp:
            json.loads(resp.read())
        print("  [텔레그램] 전송 완료")
    except Exception as e:
        print(f"  [텔레그램] 전송 실패: {e}")


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 65)
    print("  Forward Test 평가")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    if not os.path.exists(HISTORY_FILE):
        print("[오류] forward_test_history.json 없음.")
        print("       먼저 auto_research_value.py 실행 필요.")
        sys.exit(1)

    with open(HISTORY_FILE, encoding="utf-8") as f:
        history_data = json.load(f)

    records = history_data.get("records", [])
    print(f"\n[1] 추천 기록: {len(records)}건 로드")

    # 각 기록 평가
    evaluations = []
    for i, record in enumerate(records, 1):
        print(f"\n  [{i}/{len(records)}] {record.get('date')} 평가 중...")
        ev = evaluate_record(record)
        if ev:
            evaluations.append(ev)
            print(f"    완료: {ev['days_elapsed']}일 경과, {len(ev['horizons'])}개 horizon 평가")
        else:
            print(f"    스킵 (평가에 충분한 시간 미경과)")

    print(f"\n[2] 평가 완료: {len(evaluations)}건")

    # 누적 통계
    stats = calculate_cumulative_stats(evaluations)
    print(f"\n[누적 통계]")
    for h, s in stats.items():
        print(f"  {s['horizon_days']}일: 알파 {s['avg_realized_alpha_pct']:+.2f}%, "
              f"Hit Rate {s['avg_hit_rate_pct']}%, IR {s['information_ratio']}")

    # 학습 곡선
    learning_curve = calculate_learning_curve(evaluations)

    # 결과 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "total_recommendations": len(records),
        "evaluated_recommendations": len(evaluations),
        "cumulative_stats": stats,
        "learning_curve": learning_curve,
        "evaluations": evaluations,
        "warning": "🚨 시뮬레이션 평가. 자동 매매 금지.",
    }

    os.makedirs(os.path.dirname(EVAL_FILE), exist_ok=True)
    with open(EVAL_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n  결과 저장: {EVAL_FILE}")

    # 텔레그램 알림
    if stats:
        send_telegram_summary(stats, len(evaluations))

    print("\n" + "=" * 65)
    print("  ⚠️ 시뮬레이션 전용. 자동 매매 금지.")
    print("=" * 65)


if __name__ == "__main__":
    main()
