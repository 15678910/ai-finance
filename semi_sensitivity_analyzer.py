"""
반도체 민감도 분석기
====================

미국 선물/지수 1% 변동 시 한국 반도체주의 예상 변동을 회귀분석으로 계산.

분석 대상:
  미국 (X 변수):
    - NQ=F: Nasdaq 100 선물 (AI/기술주 핵심)
    - ES=F: S&P 500 선물 (광범위 위험선호)
    - ^SOX: PHLX 반도체 지수 (업종 직접)
    - ^GSPC: S&P 500 현물
    - ^IXIC: Nasdaq 현물

  한국 (Y 변수):
    - 삼성전자, SK하이닉스, 한미반도체
    - DB하이텍, LG에너지솔루션, 삼성SDI

방법: 최근 12개월 일별 수익률 OLS 회귀 → 베타 + R²

🚨 시뮬레이션. 자동 매매 금지.
"""

import os
import sys
import json
from datetime import datetime, timezone, timedelta

try:
    import yfinance as yf
    import pandas as pd
    import numpy as np
except ImportError as e:
    print(f"[오류] 라이브러리 미설치: {e}")
    sys.exit(1)

from core import send_message

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "semi_sensitivity.json")
KST = timezone(timedelta(hours=9))

# 미국 X 변수 (independent)
US_BENCHMARKS = [
    {"ticker": "NQ=F", "name": "Nasdaq 100 선물", "key": "nq", "primary": True},
    {"ticker": "ES=F", "name": "S&P 500 선물", "key": "es"},
    {"ticker": "^SOX", "name": "PHLX 반도체 지수", "key": "sox", "primary": True},
    {"ticker": "^GSPC", "name": "S&P 500", "key": "spx"},
    {"ticker": "^IXIC", "name": "Nasdaq", "key": "ixic"},
]

# 한국 Y 변수 (dependent)
KR_SEMIS = [
    {"ticker": "005930.KS", "name": "삼성전자"},
    {"ticker": "000660.KS", "name": "SK하이닉스"},
    {"ticker": "042700.KS", "name": "한미반도체"},
    {"ticker": "000990.KS", "name": "DB하이텍"},
    {"ticker": "373220.KS", "name": "LG에너지솔루션"},
    {"ticker": "006400.KS", "name": "삼성SDI"},
]


def fetch_daily_returns(ticker: str, period: str = "1y") -> "pd.Series":
    """일별 수익률 추출. 인덱스는 tz-naive 날짜로 정규화."""
    try:
        t = yf.Ticker(ticker)
        hist = t.history(period=period)
        if hist.empty or len(hist) < 30:
            return pd.Series(dtype=float)
        close = hist["Close"].dropna()
        # 타임존 제거 + 날짜만 추출 (시각 정보 폐기)
        idx = pd.DatetimeIndex(close.index)
        if idx.tz is not None:
            idx = idx.tz_localize(None)
        close.index = pd.DatetimeIndex(idx.date)
        returns = close.pct_change().dropna()
        return pd.Series(returns.values, index=returns.index, dtype=float)
    except Exception as e:
        print(f"  [실패] {ticker}: {e}")
        return pd.Series(dtype=float)


def calculate_beta(y: "pd.Series", x: "pd.Series", lag_x: int = 1) -> dict:
    """OLS 회귀로 베타 + R² 추정.

    Y(한국, t일) = α + β × X(미국, t-lag일) + ε
    β = Cov(Y,X) / Var(X)

    lag_x=1: 미국 T일 종가 변동이 한국 T+1일에 반영 (정상적 시간 차이)
    lag_x=0: 동일 거래일 (잘못된 정렬, R²이 매우 낮게 나옴)
    """
    # 미국 X를 lag일만큼 시프트 (X.shift(1) → X의 T일 값이 T+1일 인덱스로 이동)
    x_lagged = x.shift(lag_x).dropna()

    # 공통 날짜 인덱스 (inner join)
    df = pd.DataFrame({"y": y, "x": x_lagged}).dropna()
    if len(df) < 30:
        return {"beta": None, "r_squared": None, "n_samples": len(df), "alpha": None}

    y_vals = np.asarray(df["y"].to_numpy(), dtype=float)
    x_vals = np.asarray(df["x"].to_numpy(), dtype=float)

    # 베타 계산
    cov_matrix = np.cov(y_vals, x_vals)
    cov = float(cov_matrix[0, 1])
    var_x = float(np.var(x_vals, ddof=1))
    if var_x == 0:
        return {"beta": None, "r_squared": None, "n_samples": len(df), "alpha": None}

    beta = cov / var_x
    alpha = float(np.mean(y_vals) - beta * np.mean(x_vals))

    # R² (결정계수)
    y_pred = alpha + beta * x_vals
    ss_res = np.sum((y_vals - y_pred) ** 2)
    ss_tot = np.sum((y_vals - np.mean(y_vals)) ** 2)
    r_squared = 1 - (ss_res / ss_tot) if ss_tot > 0 else 0

    # 상관계수
    correlation = float(np.corrcoef(y_vals, x_vals)[0, 1])

    return {
        "beta": round(float(beta), 4),
        "alpha": round(alpha, 6),
        "r_squared": round(float(r_squared), 4),
        "correlation": round(correlation, 4),
        "n_samples": len(df),
    }


def interpret_beta(beta: float, r_squared: float) -> dict:
    """베타 해석."""
    if beta is None:
        return {"level": "데이터 부족", "color": "gray"}

    # R²이 낮으면 베타도 신뢰도 낮음
    confidence = "높음" if r_squared > 0.4 else "중간" if r_squared > 0.2 else "낮음"

    if abs(beta) > 1.5:
        level = "매우 민감"
        color = "red"
    elif abs(beta) > 1.0:
        level = "민감"
        color = "amber"
    elif abs(beta) > 0.5:
        level = "보통"
        color = "cyan"
    else:
        level = "둔감"
        color = "green"

    direction = "동조" if beta > 0 else "역행"

    return {
        "level": f"{direction} {level}",
        "confidence": confidence,
        "color": color,
    }


def predict_impact(beta: float, x_change_pct: float = 1.0) -> dict:
    """X 변수 ±1% 변동 시 예상 Y 변동."""
    if beta is None:
        return {}
    return {
        "up_1pct": round(beta * x_change_pct, 3),
        "up_3pct": round(beta * x_change_pct * 3, 3),
        "down_1pct": round(beta * -x_change_pct, 3),
        "down_3pct": round(beta * -x_change_pct * 3, 3),
    }


def main():
    print("=" * 65)
    print("  반도체 민감도 분석기 (NQ/SOX → 한국 반도체)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    # 미국 벤치마크 수익률 수집
    print("\n[미국 벤치마크 수익률 수집]")
    us_returns = {}
    for b in US_BENCHMARKS:
        print(f"  {b['name']}...")
        r = fetch_daily_returns(b["ticker"])
        if not r.empty:
            us_returns[b["key"]] = {"series": r, "info": b}
            print(f"    {len(r)}일 수집 완료")

    # 한국 반도체주 수익률
    print("\n[한국 반도체주 수익률 수집]")
    kr_returns = {}
    for k in KR_SEMIS:
        print(f"  {k['name']}...")
        r = fetch_daily_returns(k["ticker"])
        if not r.empty:
            kr_returns[k["ticker"]] = {"series": r, "info": k}
            print(f"    {len(r)}일 수집 완료")

    # 회귀 분석
    print("\n[베타 회귀 분석]")
    results = []
    for kr_ticker, kr_data in kr_returns.items():
        kr_info = kr_data["info"]
        kr_series = kr_data["series"]

        stock_result = {
            "name": kr_info["name"],
            "ticker": kr_ticker,
            "benchmarks": {},
        }

        for us_key, us_data in us_returns.items():
            us_info = us_data["info"]
            us_series = us_data["series"]

            beta_result = calculate_beta(kr_series, us_series)
            interp = interpret_beta(beta_result["beta"], beta_result["r_squared"])
            impact = predict_impact(beta_result["beta"])

            stock_result["benchmarks"][us_key] = {
                "benchmark_name": us_info["name"],
                "benchmark_ticker": us_info["ticker"],
                "primary": us_info.get("primary", False),
                **beta_result,
                "interpretation": interp,
                "impact": impact,
            }

        results.append(stock_result)

        # 콘솔 출력 (NQ 기준)
        nq = stock_result["benchmarks"].get("nq", {})
        sox = stock_result["benchmarks"].get("sox", {})
        if nq.get("beta") is not None:
            print(f"  {kr_info['name']:15s} | NQ β={nq['beta']:+.3f} R²={nq['r_squared']:.2f} | "
                  f"SOX β={sox.get('beta', 0):+.3f} R²={sox.get('r_squared', 0):.2f}")

    # 결과 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "analysis_period": "최근 1년 일별 수익률",
        "methodology": "OLS 회귀: Y(한국주) = α + β × X(미국벤치마크) + ε",
        "us_benchmarks": [{"ticker": b["ticker"], "name": b["name"], "key": b["key"]} for b in US_BENCHMARKS],
        "kr_stocks": [{"ticker": k["ticker"], "name": k["name"]} for k in KR_SEMIS],
        "results": results,
        "warning": "🚨 과거 데이터 기반 통계. 실시간 정확도 보장 안 됨. 자동 매매 금지.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    # 텔레그램 요약
    if results:
        lines = ["📊 반도체 민감도 분석 (NQ/SOX → 한국)", "=" * 30, ""]
        for r in results:
            nq = r["benchmarks"].get("nq", {})
            sox = r["benchmarks"].get("sox", {})
            if nq.get("beta") is not None:
                lines.append(f"{r['name']}")
                lines.append(f"  NQ선물 β={nq['beta']:+.2f} (R² {nq['r_squared']:.2f}) → NQ +1%면 {nq['impact']['up_1pct']:+.2f}%")
                lines.append(f"  SOX β={sox.get('beta', 0):+.2f} (R² {sox.get('r_squared', 0):.2f}) → SOX +1%면 {sox.get('impact', {}).get('up_1pct', 0):+.2f}%")
                lines.append("")
        lines.append("🚨 과거 통계. 자동 매매 금지.")
        lines.append("\n대시보드: https://15678910.github.io/ai-finance/")
        send_message("\n".join(lines))

    print("\n" + "=" * 65)
    print(f"  분석 완료: {len(results)}개 종목")
    print("  ⚠️ 시뮬레이션 / 과거 데이터 기반. 자동 매매 금지.")
    print("=" * 65)


if __name__ == "__main__":
    main()
