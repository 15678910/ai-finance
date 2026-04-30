"""
오토리서치 - 포트폴리오 가중치 진화 시스템 (시뮬레이션 전용)
========================================================

🚨 절대 규칙: 실전 매매 금지. 시뮬레이션 결과만 표시.

진화 루프:
  1. 베이스라인 가중치 로드
  2. ±10% 범위로 변형 (가설 생성)
  3. 백테스트 실행 (in-sample)
  4. 검증 데이터로 over-fitting 체크 (out-of-sample)
  5. 개선 시 채택, 아니면 폐기
  6. 시간 예산 소진까지 반복

황금 지표: Sortino Ratio (하락 변동성 기준)

사용법:
  python auto_research_portfolio.py --tickers 005930,000660,035420,035720
  python auto_research_portfolio.py --sector IT_반도체 --time-budget 300
"""

import os
import sys
import json
import time
import random
import argparse
import warnings
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta
from pathlib import Path

warnings.filterwarnings('ignore')

# 필수 라이브러리 확인
try:
    import numpy as np
    import pandas as pd
    import yfinance as yf
except ImportError as e:
    print(f"[오류] 필수 라이브러리 미설치: {e}")
    print("설치: pip install numpy pandas yfinance")
    sys.exit(1)

# ====================================================================
# 절대 규칙 (Constants)
# ====================================================================
TRADING_DAYS = 252
RISK_FREE_RATE = 0.025  # 2.5%
MIN_WEIGHT = 0.01       # 단일 종목 최소 1%
MAX_WEIGHT = 0.50       # 단일 종목 최대 50%
PERTURBATION_RANGE = 0.10  # ±10% 변형
DEFAULT_TIME_BUDGET = 300  # 5분
MAX_ITERATIONS = 200       # 시간 예산 무관 최대 시도

# In-sample / Out-of-sample 분할
IN_SAMPLE_DAYS = 252       # 1년
OUT_SAMPLE_DAYS = 126      # 6개월
TOTAL_DAYS_NEEDED = IN_SAMPLE_DAYS + OUT_SAMPLE_DAYS

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "auto_research_portfolio.json")


# ====================================================================
# Sortino Ratio 계산 (격리된 평가 함수 - 절대 수정 금지)
# ====================================================================
def calculate_sortino(returns: np.ndarray) -> float:
    """Sortino Ratio 계산. 단일 스칼라 황금 지표."""
    if len(returns) < 30:
        return -999.0

    ann_return = np.mean(returns) * TRADING_DAYS
    downside_returns = returns[returns < 0]
    if len(downside_returns) < 5:
        return -999.0
    downside_vol = np.std(downside_returns) * np.sqrt(TRADING_DAYS)
    if downside_vol < 1e-6:
        return -999.0

    return (ann_return - RISK_FREE_RATE) / downside_vol


def calculate_metrics(returns: np.ndarray) -> dict:
    """포트폴리오 성과 지표 계산 (참고용)."""
    if len(returns) < 30:
        return {}

    ann_return = np.mean(returns) * TRADING_DAYS
    ann_vol = np.std(returns) * np.sqrt(TRADING_DAYS)
    sortino = calculate_sortino(returns)
    sharpe = (ann_return - RISK_FREE_RATE) / ann_vol if ann_vol > 0 else 0

    # MDD
    cumulative = (1 + returns).cumprod()
    cummax = np.maximum.accumulate(cumulative)
    drawdown = (cumulative - cummax) / cummax
    max_dd = np.min(drawdown)

    return {
        "sortino": round(sortino, 3),
        "sharpe": round(sharpe, 3),
        "ann_return": round(ann_return * 100, 2),
        "ann_vol": round(ann_vol * 100, 2),
        "max_dd": round(max_dd * 100, 2),
    }


# ====================================================================
# 진화 루프 (Sandbox - 자유롭게 변경 가능)
# ====================================================================
def evolve_weights(base_weights: np.ndarray, n_iter: int = 1) -> np.ndarray:
    """베이스 가중치를 ±10% 범위로 변형."""
    weights = base_weights.copy()
    for _ in range(n_iter):
        # 랜덤하게 두 종목 선택해서 비중 이동
        idx_from = random.randint(0, len(weights) - 1)
        idx_to = random.randint(0, len(weights) - 1)
        if idx_from == idx_to:
            continue
        # 이동량 (현재 비중의 5~15%)
        transfer = weights[idx_from] * random.uniform(0.05, 0.15)
        weights[idx_from] -= transfer
        weights[idx_to] += transfer

    # 안전망: 최소/최대 비중 강제
    weights = np.clip(weights, MIN_WEIGHT, MAX_WEIGHT)
    # 정규화 (합 = 1.0)
    weights = weights / weights.sum()
    return weights


# ====================================================================
# 데이터 수집
# ====================================================================
def fetch_returns(tickers: list, period: str = "2y") -> pd.DataFrame:
    """yfinance로 일별 수익률 수집."""
    print(f"\n[수집] {len(tickers)}개 자산 데이터 다운로드 중...")
    price_data = {}

    for ticker in tickers:
        try:
            # KR 자동 감지 (.KS 또는 .KQ)
            yf_ticker = ticker
            if ticker.isdigit() and len(ticker) == 6:
                # KOSPI 우선 시도
                stock = yf.Ticker(f"{ticker}.KS")
                df = stock.history(period=period)
                if df.empty or len(df) < 10:
                    stock = yf.Ticker(f"{ticker}.KQ")
                    df = stock.history(period=period)
                yf_ticker = stock.ticker
            else:
                stock = yf.Ticker(ticker)
                df = stock.history(period=period)

            if df.empty or len(df) < TOTAL_DAYS_NEEDED:
                print(f"  [건너뜀] {ticker}: 데이터 부족 ({len(df)}일)")
                continue

            # 타임존 제거
            if df.index.tz is not None:
                df.index = df.index.tz_localize(None)

            price_data[ticker] = df['Close']
            print(f"  [OK] {yf_ticker}: {len(df)}일")
        except Exception as e:
            print(f"  [실패] {ticker}: {e}")

    if not price_data:
        return pd.DataFrame()

    # DataFrame 결합
    df = pd.DataFrame(price_data).dropna()
    returns = df.pct_change().dropna()
    return returns


# ====================================================================
# 백테스트
# ====================================================================
def backtest(returns_df: pd.DataFrame, weights: np.ndarray) -> np.ndarray:
    """가중치 기반 포트폴리오 일별 수익률 계산."""
    if len(weights) != len(returns_df.columns):
        raise ValueError(f"weights 길이 ({len(weights)}) != tickers ({len(returns_df.columns)})")
    return (returns_df.values * weights).sum(axis=1)


def split_in_out_sample(returns_df: pd.DataFrame) -> tuple:
    """In-sample / Out-of-sample 분할.
    - In-sample (학습): 처음 IN_SAMPLE_DAYS
    - Out-of-sample (검증): 마지막 OUT_SAMPLE_DAYS
    """
    n = len(returns_df)
    if n < TOTAL_DAYS_NEEDED:
        # 데이터 부족 시 70/30 비율
        split = int(n * 0.7)
        return returns_df.iloc[:split], returns_df.iloc[split:]

    return returns_df.iloc[:IN_SAMPLE_DAYS], returns_df.iloc[-OUT_SAMPLE_DAYS:]


# ====================================================================
# 베이스라인 가중치 (몬테카를로 균형형 결과 또는 동일 가중)
# ====================================================================
def get_baseline_weights(n_assets: int) -> np.ndarray:
    """베이스라인: 동일 가중 (1/N)."""
    return np.ones(n_assets) / n_assets


# ====================================================================
# 메인 진화 루프
# ====================================================================
def run_evolution(tickers: list, time_budget_sec: int, sector_name: str = "") -> dict:
    """오토리서치 진화 루프 실행."""
    print("=" * 65)
    print("  오토리서치 - 포트폴리오 가중치 진화")
    print(f"  시간 예산: {time_budget_sec}초")
    print(f"  종목 수: {len(tickers)}")
    print("=" * 65)

    # 데이터 수집
    returns_df = fetch_returns(tickers)
    if returns_df.empty:
        return {"error": "데이터 수집 실패"}

    print(f"\n  데이터 기간: {returns_df.index[0].date()} ~ {returns_df.index[-1].date()}")
    print(f"  총 거래일: {len(returns_df)}일")

    # In-sample / Out-of-sample 분할
    in_sample, out_sample = split_in_out_sample(returns_df)
    print(f"  In-sample: {len(in_sample)}일, Out-of-sample: {len(out_sample)}일")

    n_assets = len(returns_df.columns)
    asset_names = list(returns_df.columns)

    # 베이스라인
    baseline = get_baseline_weights(n_assets)
    baseline_returns_in = backtest(in_sample, baseline)
    baseline_returns_out = backtest(out_sample, baseline)
    baseline_metrics_in = calculate_metrics(baseline_returns_in)
    baseline_metrics_out = calculate_metrics(baseline_returns_out)
    baseline_sortino = baseline_metrics_in.get("sortino", -999)

    print(f"\n[베이스라인] In-sample Sortino: {baseline_sortino:.3f}")
    print(f"[베이스라인] Out-of-sample Sortino: {baseline_metrics_out.get('sortino', -999):.3f}")

    # 진화 루프
    print(f"\n[진화] 시작...")
    start_time = time.time()
    best_weights = baseline.copy()
    best_sortino = baseline_sortino
    best_metrics_in = baseline_metrics_in
    best_metrics_out = baseline_metrics_out
    iterations = 0
    accepted = 0
    rejected_overfit = 0

    history = [{
        "iteration": 0,
        "in_sortino": round(baseline_sortino, 3),
        "out_sortino": round(baseline_metrics_out.get('sortino', -999), 3),
        "accepted": True,
        "note": "베이스라인",
    }]

    while iterations < MAX_ITERATIONS and (time.time() - start_time) < time_budget_sec:
        iterations += 1

        # 가설 생성
        candidate = evolve_weights(best_weights, n_iter=random.randint(1, 3))

        # In-sample 백테스트
        candidate_returns_in = backtest(in_sample, candidate)
        candidate_sortino_in = calculate_sortino(candidate_returns_in)

        # 개선 체크
        if candidate_sortino_in <= best_sortino:
            continue

        # Out-of-sample 검증 (과적합 방지)
        candidate_returns_out = backtest(out_sample, candidate)
        candidate_sortino_out = calculate_sortino(candidate_returns_out)

        # 검증 통과 조건: out-of-sample Sortino > in-sample × 0.7
        # (검증 데이터에서 70% 이상 유지되어야 함)
        threshold = candidate_sortino_in * 0.7 if candidate_sortino_in > 0 else candidate_sortino_in * 1.3

        if candidate_sortino_out < threshold:
            rejected_overfit += 1
            history.append({
                "iteration": iterations,
                "in_sortino": round(candidate_sortino_in, 3),
                "out_sortino": round(candidate_sortino_out, 3),
                "accepted": False,
                "note": "과적합 (검증 실패)",
            })
            continue

        # 채택
        best_weights = candidate
        best_sortino = candidate_sortino_in
        best_metrics_in = calculate_metrics(candidate_returns_in)
        best_metrics_out = calculate_metrics(candidate_returns_out)
        accepted += 1

        history.append({
            "iteration": iterations,
            "in_sortino": round(candidate_sortino_in, 3),
            "out_sortino": round(candidate_sortino_out, 3),
            "accepted": True,
            "note": f"채택 #{accepted}",
        })

        print(f"  [{iterations:3d}] 채택! In: {candidate_sortino_in:.3f} → Out: {candidate_sortino_out:.3f}")

    elapsed = time.time() - start_time
    print(f"\n[완료] 총 시도: {iterations}회, 채택: {accepted}회, 과적합 거부: {rejected_overfit}회")
    print(f"       소요 시간: {elapsed:.1f}초")
    print(f"       베이스라인 Sortino: {baseline_sortino:.3f}")
    print(f"       진화 후 Sortino: {best_sortino:.3f}")
    if baseline_sortino != 0:
        improvement = (best_sortino - baseline_sortino) / abs(baseline_sortino) * 100
        print(f"       개선율: {improvement:+.1f}%")

    # 결과 정리
    weights_dict = {asset: round(float(w), 4) for asset, w in zip(asset_names, best_weights)}

    result = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sector": sector_name or "사용자_지정",
        "tickers": asset_names,
        "baseline_weights": {a: round(float(w), 4) for a, w in zip(asset_names, baseline)},
        "evolved_weights": weights_dict,
        "baseline_metrics": {
            "in_sample": baseline_metrics_in,
            "out_of_sample": baseline_metrics_out,
        },
        "evolved_metrics": {
            "in_sample": best_metrics_in,
            "out_of_sample": best_metrics_out,
        },
        "improvement_pct": round((best_sortino - baseline_sortino) / abs(baseline_sortino) * 100, 1) if baseline_sortino != 0 else 0,
        "iterations": iterations,
        "accepted": accepted,
        "rejected_overfit": rejected_overfit,
        "elapsed_sec": round(elapsed, 1),
        "validation_passed": best_metrics_out.get("sortino", -999) >= best_metrics_in.get("sortino", -999) * 0.7,
        "warning": "🚨 시뮬레이션 결과. 실제 매매 금지. 단순 참고용.",
        "history_summary": history[-10:],  # 최근 10개만
    }

    return result


# ====================================================================
# 텔레그램 전송 (선택)
# ====================================================================
def send_telegram(result: dict):
    """진화 결과를 텔레그램으로 알림."""
    env_path = os.path.join(CONFIG_DIR, ".env")
    bot_token = None
    chat_id = None

    # .env 파싱
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" not in line or line.startswith("#"):
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip().strip("'\"")
                if k == "TELEGRAM_FINANCE_BOT_TOKEN":
                    bot_token = v
                elif k == "TELEGRAM_FINANCE_CHAT_ID":
                    chat_id = v

    bot_token = bot_token or os.environ.get("TELEGRAM_FINANCE_BOT_TOKEN")
    chat_id = chat_id or os.environ.get("TELEGRAM_FINANCE_CHAT_ID")

    if not bot_token or not chat_id:
        print("  [텔레그램] 토큰 미설정. 전송 생략.")
        return

    # 메시지 조립
    sector = result.get("sector", "")
    base_in = result["baseline_metrics"]["in_sample"].get("sortino", "N/A")
    evolved_in = result["evolved_metrics"]["in_sample"].get("sortino", "N/A")
    evolved_out = result["evolved_metrics"]["out_of_sample"].get("sortino", "N/A")
    valid = "✅" if result.get("validation_passed") else "⚠️"

    msg_lines = [
        "🧪 오토리서치 진화 결과",
        "=" * 25,
        "",
        f"섹터: {sector}",
        f"베이스라인 Sortino: {base_in}",
        f"진화 Sortino (in): {evolved_in}",
        f"검증 Sortino (out): {evolved_out} {valid}",
        f"개선율: {result.get('improvement_pct', 0):+.1f}%",
        f"시도: {result.get('iterations', 0)}회 (채택: {result.get('accepted', 0)})",
        "",
        "🚨 실제 매매 금지 - 시뮬레이션 결과",
        "",
        f"대시보드: https://15678910.github.io/ai-finance/",
    ]

    message = "\n".join(msg_lines)

    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        data = urllib.parse.urlencode({"chat_id": chat_id, "text": message}).encode()
        req = urllib.request.Request(url, data=data, method="POST")
        with urllib.request.urlopen(req, timeout=10) as resp:
            json.loads(resp.read())
        print("  [텔레그램] 전송 완료")
    except Exception as e:
        print(f"  [텔레그램] 전송 실패: {e}")


# ====================================================================
# 결과 저장
# ====================================================================
def save_result(result: dict):
    """docs/auto_research_portfolio.json에 저장."""
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

    # 누적 기록 (최근 30개 유지)
    history = []
    if os.path.exists(OUTPUT_FILE):
        try:
            with open(OUTPUT_FILE, encoding="utf-8") as f:
                old = json.load(f)
                history = old.get("evolution_history", [])
        except Exception:
            pass

    # 현재 결과를 히스토리에 추가
    history_entry = {
        "date": result.get("generated_at"),
        "sector": result.get("sector"),
        "baseline_sortino": result["baseline_metrics"]["in_sample"].get("sortino"),
        "evolved_sortino": result["evolved_metrics"]["in_sample"].get("sortino"),
        "validation_sortino": result["evolved_metrics"]["out_of_sample"].get("sortino"),
        "improvement_pct": result.get("improvement_pct"),
        "iterations": result.get("iterations"),
    }
    history.append(history_entry)
    history = history[-30:]  # 최근 30개

    result["evolution_history"] = history

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"\n  결과 저장: {OUTPUT_FILE}")


# ====================================================================
# 메인
# ====================================================================
def main():
    parser = argparse.ArgumentParser(description="오토리서치 - 포트폴리오 가중치 진화")
    parser.add_argument("--tickers", default="005930,000660,035420,035720",
                        help="종목 코드 (쉼표 구분, 기본: IT/반도체)")
    parser.add_argument("--sector", default="IT_반도체", help="섹터명 (라벨)")
    parser.add_argument("--time-budget", type=int, default=DEFAULT_TIME_BUDGET,
                        help=f"시간 예산 (초, 기본 {DEFAULT_TIME_BUDGET}초)")
    parser.add_argument("--no-telegram", action="store_true", help="텔레그램 전송 생략")
    args = parser.parse_args()

    tickers = [t.strip() for t in args.tickers.split(",") if t.strip()]

    # 진화 실행
    result = run_evolution(tickers, args.time_budget, args.sector)

    if "error" in result:
        print(f"[오류] {result['error']}")
        sys.exit(1)

    # 저장
    save_result(result)

    # 텔레그램
    if not args.no_telegram:
        send_telegram(result)

    print("\n" + "=" * 65)
    print("  ⚠️  중요: 본 결과는 시뮬레이션 전용입니다.")
    print("  ⚠️  실제 매매 시 사용자 본인의 판단 필요.")
    print("  ⚠️  자동 매매 절대 불가.")
    print("=" * 65)


if __name__ == "__main__":
    main()
