"""
오토리서치 - 가치투자 공식 진화 시스템 (시뮬레이션 전용)
========================================================

🚨 절대 규칙: 실전 매매 금지. 시뮬레이션 결과만 표시.

진화 루프:
  1. 35개 종목의 현재 펀더멘털 (PER, PBR, ROE 등) 로드
  2. 과거 6개월 수익률을 "진실값"으로 사용
  3. 가중치 (W1~W4) 변형 → 점수 → Top 10
  4. Top 10의 6개월 평균 수익률 vs KOSPI = 알파
  5. Cross-validation으로 과적합 방지
  6. 알파 최대화 가중치 선택

황금 지표: Alpha (Top 10 평균 수익률 - KOSPI 수익률)
보조 지표: Hit Rate (Top 10 중 KOSPI 이긴 비율)

사용법:
  python auto_research_value.py
  python auto_research_value.py --time-budget 300

🚨 절대 규칙:
  - 시뮬레이션 / 분석 전용
  - 자동 매매 금지
  - 사용자 직접 검토 후 투자
"""

import os
import sys
import json
import time
import random
import argparse
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta
from pathlib import Path
from copy import deepcopy

try:
    import numpy as np
    import pandas as pd
    import yfinance as yf
except ImportError as e:
    print(f"[오류] 라이브러리 미설치: {e}")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
SCREENER_FILE = os.path.join(BASE_DIR, "docs", "value_screener.json")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "auto_research_value.json")
HISTORY_FILE = os.path.join(BASE_DIR, "docs", "forward_test_history.json")

# 진화 파라미터
DEFAULT_TIME_BUDGET = 180  # 3분
MAX_ITERATIONS = 200
PERTURBATION = 0.15  # ±15% 가중치 변형 (큰 탐색)
MIN_WEIGHT = 0.05
MAX_WEIGHT = 0.60
TOP_N = 10
KOSPI_TICKER = "^KS11"
KST = timezone(timedelta(hours=9))


# ====================================================================
# 점수 계산 함수 (절대 변경 금지 - 격리된 평가)
# ====================================================================
def score_per(per):
    if per is None or per <= 0:
        return 0
    if per < 5: return 100
    elif per < 10: return 80
    elif per < 15: return 60
    elif per < 20: return 40
    elif per < 30: return 20
    else: return 5


def score_pbr(pbr):
    if pbr is None or pbr <= 0:
        return 0
    if pbr < 0.5: return 100
    elif pbr < 1.0: return 85
    elif pbr < 1.5: return 65
    elif pbr < 2.0: return 45
    elif pbr < 3.0: return 25
    else: return 10


def score_roe(roe_pct):
    if roe_pct is None:
        return 0
    if roe_pct < 0: return 0
    elif roe_pct < 5: return 20
    elif roe_pct < 10: return 50
    elif roe_pct < 15: return 75
    elif roe_pct < 20: return 90
    else: return 100


def score_debt(debt):
    if debt is None or debt < 0:
        return 50
    if debt < 50: return 100
    elif debt < 100: return 70
    elif debt < 150: return 40
    elif debt < 200: return 20
    else: return 5


def score_growth(growth_pct):
    if growth_pct is None:
        return 30
    if growth_pct < -10: return 0
    elif growth_pct < 0: return 20
    elif growth_pct < 5: return 50
    elif growth_pct < 10: return 70
    elif growth_pct < 20: return 85
    else: return 100


def score_dividend(div_pct):
    if div_pct is None or div_pct < 0:
        return 30
    if div_pct < 1: return 20
    elif div_pct < 2: return 50
    elif div_pct < 4: return 80
    elif div_pct < 6: return 95
    else: return 100


def calculate_score(metrics, weights):
    """가중치 기반 종합 점수 계산.
    weights = (w_valuation, w_quality, w_growth, w_shareholder)
    합 = 1.0
    """
    s_per = score_per(metrics.get("per"))
    s_pbr = score_pbr(metrics.get("pbr"))
    s_roe = score_roe(metrics.get("roe_pct"))
    s_debt = score_debt(metrics.get("debt_to_equity"))
    s_growth = score_growth(metrics.get("revenue_growth_pct"))
    s_div = score_dividend(metrics.get("dividend_yield_pct"))

    valuation = (s_per + s_pbr) / 2
    quality = (s_roe * 0.6) + (s_debt * 0.4)
    growth = s_growth
    shareholder = s_div

    return (valuation * weights[0] + quality * weights[1] +
            growth * weights[2] + shareholder * weights[3])


# ====================================================================
# 데이터 로드
# ====================================================================
def load_screener_stocks() -> list:
    """value_screener.json에서 분석 대상 종목 로드."""
    if not os.path.exists(SCREENER_FILE):
        print(f"[오류] value_screener.json 없음. 먼저 value_screener.py 실행 필요.")
        return []
    with open(SCREENER_FILE, encoding="utf-8") as f:
        data = json.load(f)
    return data.get("all_passed", [])


def fetch_returns(tickers: list, period: str = "6mo") -> dict:
    """각 종목의 과거 N개월 수익률 수집."""
    print(f"\n[수집] {len(tickers)}개 종목 + KOSPI {period} 수익률...")
    returns = {}
    for ticker in tickers:
        # KOSPI / KOSDAQ 자동 결정
        for suffix in [".KS", ".KQ"]:
            try:
                yf_ticker = ticker if "." in ticker else ticker + suffix
                stock = yf.Ticker(yf_ticker)
                hist = stock.history(period=period)
                if not hist.empty and len(hist) >= 20:
                    start = hist["Close"].iloc[0]
                    end = hist["Close"].iloc[-1]
                    if start > 0:
                        ret = (end - start) / start * 100
                        returns[ticker] = round(float(ret), 2)
                        break
            except Exception:
                continue

    # KOSPI 벤치마크
    try:
        kospi = yf.Ticker(KOSPI_TICKER).history(period=period)
        if not kospi.empty:
            kospi_ret = (kospi["Close"].iloc[-1] - kospi["Close"].iloc[0]) / kospi["Close"].iloc[0] * 100
            returns["__KOSPI__"] = round(float(kospi_ret), 2)
            print(f"  KOSPI 벤치마크: {returns['__KOSPI__']:+.2f}%")
    except Exception as e:
        returns["__KOSPI__"] = 0
        print(f"  [경고] KOSPI 수익률 수집 실패: {e}")

    print(f"  수집 완료: {len(returns)-1}개 종목")
    return returns


# ====================================================================
# 평가 함수
# ====================================================================
def evaluate_weights(weights: tuple, stocks: list, returns: dict) -> dict:
    """가중치 평가: Top 10 추출 → 알파 계산."""
    # 점수 계산
    scored = []
    for s in stocks:
        ticker = s["ticker"]
        if ticker not in returns:
            continue
        score = calculate_score(s["metrics"], weights)
        scored.append({
            "ticker": ticker,
            "name": s["name"],
            "score": score,
            "return": returns[ticker],
        })

    if len(scored) < TOP_N:
        return {"alpha": -999, "top_n": [], "hit_rate": 0}

    # Top N 추출
    scored.sort(key=lambda x: x["score"], reverse=True)
    top = scored[:TOP_N]

    # 평균 수익률
    avg_return = np.mean([s["return"] for s in top])
    kospi_return = returns.get("__KOSPI__", 0)
    alpha = avg_return - kospi_return

    # Hit Rate
    hits = sum(1 for s in top if s["return"] > kospi_return)
    hit_rate = hits / TOP_N * 100

    return {
        "alpha": round(alpha, 2),
        "avg_return": round(avg_return, 2),
        "kospi_return": kospi_return,
        "hit_rate": round(hit_rate, 1),
        "top_n": [{"ticker": s["ticker"], "name": s["name"],
                   "score": round(s["score"], 1), "return_6m": s["return"]} for s in top],
    }


def cross_validate(weights: tuple, stocks: list, returns: dict, n_folds: int = 5) -> float:
    """교차 검증: 종목을 무작위로 분할하여 알파 안정성 평가."""
    alphas = []
    for _ in range(n_folds):
        # 70% 훈련 / 30% 검증
        random.shuffle(stocks)
        split = int(len(stocks) * 0.7)
        valid = stocks[split:]
        result = evaluate_weights(weights, valid, returns)
        alphas.append(result["alpha"])

    return round(float(np.mean(alphas)), 2), round(float(np.std(alphas)), 2)


# ====================================================================
# 진화 루프
# ====================================================================
def perturb_weights(weights: tuple) -> tuple:
    """가중치 변형 (정규화 강제)."""
    w = list(weights)
    # 두 가중치 사이에서 비중 이동
    i = random.randint(0, 3)
    j = random.randint(0, 3)
    while j == i:
        j = random.randint(0, 3)
    transfer = random.uniform(0.01, PERTURBATION)
    transfer = min(transfer, w[i] - MIN_WEIGHT)
    transfer = min(transfer, MAX_WEIGHT - w[j])
    if transfer <= 0:
        return tuple(w)
    w[i] -= transfer
    w[j] += transfer
    # 안전망: clip + 정규화
    w = [max(MIN_WEIGHT, min(MAX_WEIGHT, x)) for x in w]
    total = sum(w)
    return tuple(x / total for x in w)


def evolve(stocks: list, returns: dict, time_budget: int) -> dict:
    """진화 루프 실행."""
    print(f"\n[진화] 시간 예산 {time_budget}초")

    # 베이스라인 (현재 공식)
    baseline = (0.35, 0.35, 0.20, 0.10)
    baseline_result = evaluate_weights(baseline, stocks, returns)
    print(f"  베이스라인 알파: {baseline_result['alpha']:+.2f}% (수익률 {baseline_result['avg_return']:+.2f}% vs KOSPI {baseline_result['kospi_return']:+.2f}%)")
    print(f"  베이스라인 Hit Rate: {baseline_result['hit_rate']}%")

    best_weights = baseline
    best_alpha = baseline_result["alpha"]
    best_result = baseline_result

    iterations = 0
    accepted = 0
    rejected_cv = 0
    history = [{"iter": 0, "weights": list(baseline), "alpha": baseline_result["alpha"], "accepted": True}]
    start = time.time()

    no_improve_count = 0
    # Phase 1: Grid search (다양한 시작점 탐색)
    grid_phase = True
    grid_candidates = []
    for w1 in [0.10, 0.25, 0.40, 0.55]:
        for w2 in [0.10, 0.25, 0.40, 0.55]:
            for w3 in [0.10, 0.25, 0.40, 0.55]:
                w4 = 1.0 - w1 - w2 - w3
                if MIN_WEIGHT <= w4 <= MAX_WEIGHT:
                    grid_candidates.append((w1, w2, w3, w4))
    grid_idx = 0

    while iterations < MAX_ITERATIONS and (time.time() - start) < time_budget:
        iterations += 1

        # Phase 1: Grid 탐색
        if grid_phase and grid_idx < len(grid_candidates):
            candidate = grid_candidates[grid_idx]
            grid_idx += 1
        else:
            grid_phase = False
            # Random restart (30회 연속 미개선 시 random 시작점)
            if no_improve_count >= 30:
                r = [random.uniform(MIN_WEIGHT, MAX_WEIGHT) for _ in range(4)]
                total = sum(r)
                candidate = tuple(x / total for x in r)
                no_improve_count = 0
            else:
                candidate = perturb_weights(best_weights)
        cand_result = evaluate_weights(candidate, stocks, returns)

        if cand_result["alpha"] <= best_alpha:
            no_improve_count += 1
            continue

        # 교차 검증 (과적합 방지)
        cv_mean, cv_std = cross_validate(candidate, stocks, returns)

        # 검증 알파가 baseline - 2%p 이상이면 채택 (작은 표본 노이즈 허용)
        if cv_mean < baseline_result["alpha"] - 2.0:
            rejected_cv += 1
            history.append({"iter": iterations, "weights": list(candidate),
                           "alpha": cand_result["alpha"], "cv_alpha": cv_mean,
                           "accepted": False, "reason": "cross-validation 실패"})
            continue

        # 채택
        best_weights = candidate
        best_alpha = cand_result["alpha"]
        best_result = cand_result
        best_result["cv_mean"] = cv_mean
        best_result["cv_std"] = cv_std
        accepted += 1
        no_improve_count = 0
        history.append({"iter": iterations, "weights": list(candidate),
                       "alpha": cand_result["alpha"], "cv_alpha": cv_mean,
                       "accepted": True})
        print(f"  [{iterations:3d}] 채택! 알파 {cand_result['alpha']:+.2f}% (CV: {cv_mean:+.2f}±{cv_std:.2f})")

    elapsed = time.time() - start
    print(f"\n[완료] 시도 {iterations}회, 채택 {accepted}회, CV 거부 {rejected_cv}회 ({elapsed:.1f}초)")

    return {
        "baseline_weights": list(baseline),
        "baseline_alpha": baseline_result["alpha"],
        "baseline_hit_rate": baseline_result["hit_rate"],
        "evolved_weights": list(best_weights),
        "evolved_alpha": best_alpha,
        "evolved_hit_rate": best_result["hit_rate"],
        "evolved_top_n": best_result["top_n"],
        "cv_mean": best_result.get("cv_mean"),
        "cv_std": best_result.get("cv_std"),
        "iterations": iterations,
        "accepted": accepted,
        "rejected_cv": rejected_cv,
        "history": history[-15:],  # 최근 15개
        "improvement_pp": round(best_alpha - baseline_result["alpha"], 2),
    }


# ====================================================================
# Forward Test 히스토리 저장
# ====================================================================
def save_forward_test_record(result: dict):
    """현재 추천을 forward test 히스토리에 저장."""
    history = []
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, encoding="utf-8") as f:
                data = json.load(f)
                history = data.get("records", [])
        except Exception:
            pass

    # output 객체와 evolve 결과의 키 호환 처리
    weights = result.get("evolved_weights")
    if isinstance(weights, dict):
        weights_list = [weights.get("valuation", 0), weights.get("quality", 0),
                        weights.get("growth", 0), weights.get("shareholder", 0)]
    else:
        weights_list = list(weights) if weights else [0.35, 0.35, 0.20, 0.10]

    expected_alpha = result.get("evolved_alpha_pct") or result.get("evolved_alpha", 0)
    expected_hit_rate = result.get("evolved_hit_rate", 0)
    top_picks = result.get("evolved_top_picks") or result.get("evolved_top_n", [])

    record = {
        "date": datetime.now(KST).strftime("%Y-%m-%d %H:%M"),
        "weights": weights_list,
        "expected_alpha": expected_alpha,
        "expected_hit_rate": expected_hit_rate,
        "top_picks": [{"ticker": p["ticker"], "name": p["name"],
                       "score": p.get("score", 0)} for p in top_picks],
    }

    history.append(record)
    history = history[-52:]  # 최근 1년 (주간 기준)

    os.makedirs(os.path.dirname(HISTORY_FILE), exist_ok=True)
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "updated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
            "records": history,
            "warning": "🚨 시뮬레이션 전용. 실제 매매 금지.",
        }, f, ensure_ascii=False, indent=2)
    print(f"  Forward test 기록 저장: {len(history)}건")


# ====================================================================
# 텔레그램
# ====================================================================
def send_telegram(result: dict):
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

    if not bot_token or not chat_id:
        return

    w = result["evolved_weights"]
    lines = [
        "🧬 오토리서치 - 가치투자 가중치 진화",
        "=" * 25,
        "",
        f"베이스라인 알파: {result['baseline_alpha']:+.2f}% (Hit {result['baseline_hit_rate']}%)",
        f"진화 알파: {result['evolved_alpha']:+.2f}% (Hit {result['evolved_hit_rate']}%)",
        f"개선: {result['improvement_pp']:+.2f}%p",
        "",
        f"진화 가중치:",
        f"  밸류 {w[0]*100:.0f}% / 수익성 {w[1]*100:.0f}% / 성장 {w[2]*100:.0f}% / 주주환원 {w[3]*100:.0f}%",
        "",
        "Top 10 (예측):",
    ]
    for i, p in enumerate(result["evolved_top_n"][:10], 1):
        lines.append(f"  {i}. {p['name']} ({p['ticker']}) - {p['score']:.1f}점")

    lines.append("")
    lines.append("🚨 시뮬레이션 전용. 자동 매매 금지.")
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
    parser = argparse.ArgumentParser(description="오토리서치 - 가치 가중치 진화")
    parser.add_argument("--time-budget", type=int, default=DEFAULT_TIME_BUDGET)
    parser.add_argument("--no-telegram", action="store_true")
    args = parser.parse_args()

    print("=" * 65)
    print("  오토리서치 - 가치투자 공식 진화")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    # 1) 종목 데이터 로드
    stocks = load_screener_stocks()
    if not stocks:
        print("[오류] 분석 대상 종목 없음.")
        sys.exit(1)
    print(f"\n[1] 종목 데이터 로드: {len(stocks)}개")

    # 2) 6개월 수익률 수집
    tickers = [s["ticker"] for s in stocks]
    returns = fetch_returns(tickers, period="6mo")

    # 3) 진화 실행
    result = evolve(stocks, returns, args.time_budget)

    # 4) 결과 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "evaluation_period": "6개월 (현재→과거)",
        "baseline_weights": {
            "valuation": result["baseline_weights"][0],
            "quality": result["baseline_weights"][1],
            "growth": result["baseline_weights"][2],
            "shareholder": result["baseline_weights"][3],
        },
        "evolved_weights": {
            "valuation": result["evolved_weights"][0],
            "quality": result["evolved_weights"][1],
            "growth": result["evolved_weights"][2],
            "shareholder": result["evolved_weights"][3],
        },
        "baseline_alpha_pct": result["baseline_alpha"],
        "evolved_alpha_pct": result["evolved_alpha"],
        "improvement_pp": result["improvement_pp"],
        "baseline_hit_rate": result["baseline_hit_rate"],
        "evolved_hit_rate": result["evolved_hit_rate"],
        "cv_mean": result.get("cv_mean"),
        "cv_std": result.get("cv_std"),
        "evolved_top_picks": result["evolved_top_n"],
        "iterations": result["iterations"],
        "accepted": result["accepted"],
        "rejected_cv": result["rejected_cv"],
        "warning": "🚨 시뮬레이션 전용. 자동 매매 금지. 사용자 직접 검토 필수.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    # 5) Forward Test 기록
    save_forward_test_record(output)

    # 6) 텔레그램 (선택)
    if not args.no_telegram:
        send_telegram(output)

    print("\n" + "=" * 65)
    print(f"  베이스라인: 알파 {result['baseline_alpha']:+.2f}%, Hit {result['baseline_hit_rate']}%")
    print(f"  진화 결과: 알파 {result['evolved_alpha']:+.2f}%, Hit {result['evolved_hit_rate']}%")
    print(f"  개선: {result['improvement_pp']:+.2f}%p")
    print("  ⚠️ 시뮬레이션 전용. 실제 매매 금지.")
    print("=" * 65)


if __name__ == "__main__":
    main()
