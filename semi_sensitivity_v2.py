"""
반도체 민감도 분석기 v2 — 다변량 회귀 + 이벤트 더미 + 차트
=============================================================

v1 대비 강화:
  1. 다변량 회귀: NQ + SOX + USDKRW + UST10Y + KOSPI 동시 투입
  2. 어닝 발표일 ±2일 더미 변수 (일시적 민감도 변화 통제)
  3. matplotlib PNG 차트 (정규화 가격, 누적수익률)
  4. v1 단변량 R² → v2 다변량 adj R² 향상 비교

방법:
  Y(한국, t) = α + β_NQ·NQ(t-1) + β_SOX·SOX(t-1) + β_FX·USDKRW(t)
             + β_UST·ΔUST10Y(t-1) + β_KOSPI·KOSPI(t) + γ·EarningsDummy + ε

  - 미국 변수: T-1일 lag (미국 종가 → 한국 다음 날 개장)
  - 한국 변수 (USDKRW, KOSPI): T일 동일 거래일 (한국 시간 내 트레이딩)
  - UST10Y: ^TNX는 yield 자체이므로 절대 변화(diff)

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
    import matplotlib
    matplotlib.use("Agg")  # 헤드리스 백엔드
    import matplotlib.pyplot as plt
except ImportError as e:
    print(f"[오류] 라이브러리 미설치: {e}")
    sys.exit(1)

from core import send_message

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "semi_sensitivity_v2.json")
CHARTS_DIR = os.path.join(BASE_DIR, "docs", "charts")
KST = timezone(timedelta(hours=9))

# X 변수 정의 (미국 = T-1일 lag, 한국 = T일)
US_VARIABLES = [
    {"ticker": "NQ=F",  "name": "Nasdaq 100 선물", "key": "nq",     "lag": 1, "is_yield": False},
    {"ticker": "^SOX",  "name": "PHLX 반도체",     "key": "sox",    "lag": 1, "is_yield": False},
    {"ticker": "^TNX",  "name": "US 10Y 금리",     "key": "ust10y", "lag": 1, "is_yield": True},
]
KR_VARIABLES = [
    {"ticker": "KRW=X", "name": "USD/KRW",         "key": "usdkrw", "lag": 0, "is_yield": False},
    {"ticker": "^KS11", "name": "KOSPI",           "key": "kospi",  "lag": 0, "is_yield": False},
]

# Y 변수
KR_SEMIS = [
    {"ticker": "005930.KS", "name": "삼성전자"},
    {"ticker": "000660.KS", "name": "SK하이닉스"},
    {"ticker": "042700.KS", "name": "한미반도체"},
    {"ticker": "000990.KS", "name": "DB하이텍"},
    {"ticker": "373220.KS", "name": "LG에너지솔루션"},
    {"ticker": "006400.KS", "name": "삼성SDI"},
]


def fetch_series(ticker: str, period: str = "1y", is_yield: bool = False) -> "pd.Series":
    """일별 수익률 (또는 yield 변화) 추출. tz-naive 날짜 인덱스."""
    try:
        t = yf.Ticker(ticker)
        hist = t.history(period=period)
        if hist.empty or len(hist) < 30:
            return pd.Series(dtype=float)
        close = hist["Close"].dropna()
        idx = pd.DatetimeIndex(close.index)
        if idx.tz is not None:
            idx = idx.tz_localize(None)
        close.index = pd.DatetimeIndex(idx.date)
        if is_yield:
            change = close.diff().dropna()  # bp 단위 절대 변화
        else:
            change = close.pct_change().dropna()
        return pd.Series(change.to_numpy(), index=change.index, dtype=float)
    except Exception as e:
        print(f"  [실패] {ticker}: {e}")
        return pd.Series(dtype=float)


def fetch_price(ticker: str, period: str = "1y") -> "pd.Series":
    """원본 종가 시리즈 (차트용)."""
    try:
        t = yf.Ticker(ticker)
        hist = t.history(period=period)
        if hist.empty:
            return pd.Series(dtype=float)
        close = hist["Close"].dropna()
        idx = pd.DatetimeIndex(close.index)
        if idx.tz is not None:
            idx = idx.tz_localize(None)
        close.index = pd.DatetimeIndex(idx.date)
        return pd.Series(close.to_numpy(), index=close.index, dtype=float)
    except Exception:
        return pd.Series(dtype=float)


def get_earnings_dates(ticker: str) -> set:
    """yfinance에서 과거 어닝 발표일 추출."""
    try:
        t = yf.Ticker(ticker)
        ed = t.earnings_dates
        if ed is None or ed.empty:
            return set()
        dates = set()
        for d in ed.index:
            try:
                ts = pd.Timestamp(d)
                dates.add(ts.date())
            except Exception:
                continue
        return dates
    except Exception:
        return set()


def make_event_dummy(index, event_dates, window: int = 2) -> "pd.Series":
    """어닝 ±window일에 1, 아니면 0."""
    dummy = pd.Series(0.0, index=index)
    if not event_dates:
        return dummy
    event_ts = [pd.Timestamp(d) for d in event_dates]
    for d in index:
        d_ts = pd.Timestamp(d)
        for ed in event_ts:
            if abs((d_ts - ed).days) <= window:
                dummy.loc[d] = 1.0
                break
    return dummy


def multivariate_ols(y: "pd.Series", X: "pd.DataFrame") -> dict | None:
    """다변량 OLS: y = α + β·X + ε. numpy.linalg.lstsq 사용."""
    df = pd.concat([y.rename("y"), X], axis=1).dropna()
    if len(df) < 30:
        return None

    y_vals = np.asarray(df["y"].to_numpy(), dtype=float)
    X_cols = [c for c in df.columns if c != "y"]
    X_vals = np.asarray(df[X_cols].to_numpy(), dtype=float)
    n, k = X_vals.shape

    # 상수항 추가
    X_design = np.column_stack([np.ones(n), X_vals])

    try:
        coeffs, _, _, _ = np.linalg.lstsq(X_design, y_vals, rcond=None)
    except np.linalg.LinAlgError:
        return None

    alpha = float(coeffs[0])
    betas = {col: round(float(coeffs[i + 1]), 4) for i, col in enumerate(X_cols)}

    y_pred = X_design @ coeffs
    ss_res = float(np.sum((y_vals - y_pred) ** 2))
    ss_tot = float(np.sum((y_vals - np.mean(y_vals)) ** 2))
    r_squared = 1 - ss_res / ss_tot if ss_tot > 0 else 0.0

    # Adjusted R²
    if n - k - 1 > 0:
        adj_r2 = 1 - (1 - r_squared) * (n - 1) / (n - k - 1)
    else:
        adj_r2 = r_squared

    # 표준 오차 추정 (간단)
    residual_var = ss_res / max(n - k - 1, 1)
    try:
        xtx_inv = np.linalg.inv(X_design.T @ X_design)
        se_diag = np.sqrt(np.maximum(np.diag(xtx_inv) * residual_var, 0))
        # t-stat for each predictor (excluding alpha)
        t_stats = {col: round(float(coeffs[i + 1] / se_diag[i + 1]), 2) if se_diag[i + 1] > 0 else None
                   for i, col in enumerate(X_cols)}
    except np.linalg.LinAlgError:
        t_stats = {col: None for col in X_cols}

    return {
        "alpha": round(alpha, 6),
        "betas": betas,
        "t_stats": t_stats,
        "r_squared": round(float(r_squared), 4),
        "adj_r_squared": round(float(adj_r2), 4),
        "n_samples": n,
        "n_predictors": k,
    }


def generate_charts(price_data: dict, return_data: dict) -> list:
    """정규화 가격 + 누적수익률 PNG 생성. 영문 라벨 (matplotlib 한글 폰트 의존 회피)."""
    os.makedirs(CHARTS_DIR, exist_ok=True)
    saved = []

    # 1. 정규화 가격 (시작값 = 100)
    fig, ax = plt.subplots(figsize=(12, 6))
    for label, series in price_data.items():
        if series.empty:
            continue
        normalized = series / series.iloc[0] * 100
        ax.plot(normalized.index, normalized.values, label=label, linewidth=1.4)
    ax.set_title("Normalized Prices (start=100) - Last 12 months", fontsize=13)
    ax.set_ylabel("Normalized Price")
    ax.grid(True, alpha=0.3)
    ax.legend(loc="best", framealpha=0.9, fontsize=9)
    ax.axhline(100, color="gray", linewidth=0.5, linestyle="--")
    fig.tight_layout()
    path1 = os.path.join(CHARTS_DIR, "semi_normalized.png")
    fig.savefig(path1, dpi=100, bbox_inches="tight")
    plt.close(fig)
    saved.append(path1)

    # 2. 누적수익률 (%)
    fig, ax = plt.subplots(figsize=(12, 6))
    for label, series in return_data.items():
        if series.empty:
            continue
        cumret = ((1 + series).cumprod() - 1) * 100
        ax.plot(cumret.index, cumret.values, label=label, linewidth=1.4)
    ax.set_title("Cumulative Returns (%) - Last 12 months", fontsize=13)
    ax.set_ylabel("Cumulative Return (%)")
    ax.grid(True, alpha=0.3)
    ax.legend(loc="best", framealpha=0.9, fontsize=9)
    ax.axhline(0, color="gray", linewidth=0.5)
    fig.tight_layout()
    path2 = os.path.join(CHARTS_DIR, "semi_cumulative.png")
    fig.savefig(path2, dpi=100, bbox_inches="tight")
    plt.close(fig)
    saved.append(path2)

    return saved


def romanize(name: str) -> str:
    """한글 종목명 → 영문 라벨 (차트용)."""
    mapping = {
        "삼성전자": "Samsung Elec",
        "SK하이닉스": "SK Hynix",
        "한미반도체": "Hanmi Semi",
        "DB하이텍": "DB HiTek",
        "LG에너지솔루션": "LG Energy",
        "삼성SDI": "Samsung SDI",
        "Nasdaq 100 선물": "NQ Futures",
        "PHLX 반도체": "PHLX SOX",
        "USD/KRW": "USD/KRW",
        "KOSPI": "KOSPI",
        "US 10Y 금리": "UST 10Y",
    }
    return mapping.get(name, name)


def main():
    print("=" * 72)
    print("  반도체 민감도 분석기 v2 (다변량 + 이벤트 더미 + 차트)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 72)

    # X 변수 시리즈 (lag 적용)
    print("\n[X 변수 수집]")
    X_series = {}
    X_prices = {}
    for var in US_VARIABLES + KR_VARIABLES:
        print(f"  {var['name']} ({var['ticker']}, lag={var['lag']}, yield={var['is_yield']})...")
        s = fetch_series(var["ticker"], is_yield=var["is_yield"])
        if not s.empty:
            if var["lag"] > 0:
                s = s.shift(var["lag"]).dropna()
            X_series[var["key"]] = s
            print(f"    {len(s)}일 수집")
        # 차트용 원본 가격 (yield 제외)
        if not var["is_yield"]:
            p = fetch_price(var["ticker"])
            if not p.empty:
                X_prices[var["name"]] = p

    # Y 변수
    print("\n[Y 변수 (한국 반도체) 수집]")
    Y_series = {}
    Y_prices = {}
    for stock in KR_SEMIS:
        print(f"  {stock['name']}...")
        s = fetch_series(stock["ticker"])
        if not s.empty:
            Y_series[stock["ticker"]] = s
            p = fetch_price(stock["ticker"])
            if not p.empty:
                Y_prices[stock["name"]] = p

    # X DataFrame
    X_df = pd.DataFrame(X_series)
    print(f"\n  X 매트릭스: {X_df.shape[0]}행 × {X_df.shape[1]}열 ({list(X_df.columns)})")

    # 회귀 분석
    print("\n[다변량 회귀 분석]")
    results = []
    for stock in KR_SEMIS:
        ticker = stock["ticker"]
        if ticker not in Y_series:
            continue
        y = Y_series[ticker]

        # (1) 다변량 (이벤트 더미 없음)
        mv = multivariate_ols(y, X_df)

        # (2) 어닝 발표일 ±2일 더미 추가
        earnings = get_earnings_dates(ticker)
        if earnings and mv is not None:
            dummy = make_event_dummy(X_df.index, earnings, window=2)
            X_with_dummy = X_df.copy()
            X_with_dummy["earnings_window"] = dummy
            mv_event = multivariate_ols(y, X_with_dummy)
        else:
            mv_event = None

        stock_result = {
            "name": stock["name"],
            "ticker": ticker,
            "multivariate": mv,
            "multivariate_with_event": mv_event,
            "earnings_dates_count": len(earnings),
        }
        results.append(stock_result)

        if mv:
            b = mv["betas"]
            tag = ""
            if mv_event:
                delta_r2 = mv_event["r_squared"] - mv["r_squared"]
                tag = f" | +event ΔR²={delta_r2:+.3f}"
            print(f"  {stock['name']:14s} R²={mv['r_squared']:.3f} adjR²={mv['adj_r_squared']:.3f}{tag}")
            print(f"     NQ={b.get('nq', 0):+.2f}  SOX={b.get('sox', 0):+.2f}  "
                  f"USDKRW={b.get('usdkrw', 0):+.2f}  UST10Y={b.get('ust10y', 0):+.3f}  "
                  f"KOSPI={b.get('kospi', 0):+.2f}")

    # 차트 생성
    print("\n[PNG 차트 생성]")
    chart_prices = {}
    if "Nasdaq 100 선물" in X_prices:
        chart_prices[romanize("Nasdaq 100 선물")] = X_prices["Nasdaq 100 선물"]
    if "PHLX 반도체" in X_prices:
        chart_prices[romanize("PHLX 반도체")] = X_prices["PHLX 반도체"]
    for name in ["삼성전자", "SK하이닉스", "한미반도체"]:
        if name in Y_prices:
            chart_prices[romanize(name)] = Y_prices[name]

    chart_returns = {}
    if "nq" in X_series:
        chart_returns["NQ Futures"] = X_series["nq"]
    if "sox" in X_series:
        chart_returns["PHLX SOX"] = X_series["sox"]
    for stock in KR_SEMIS[:3]:
        if stock["ticker"] in Y_series:
            chart_returns[romanize(stock["name"])] = Y_series[stock["ticker"]]

    try:
        saved = generate_charts(chart_prices, chart_returns)
        print(f"  저장: {[os.path.basename(p) for p in saved]}")
    except Exception as e:
        print(f"  [차트 실패] {e}")

    # 결과 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "analysis_period": "최근 1년 일별 수익률",
        "methodology": "다변량 OLS: Y(t) = α + β_NQ·NQ(t-1) + β_SOX·SOX(t-1) + β_FX·USDKRW(t) + β_UST·ΔUST10Y(t-1) + β_KOSPI·KOSPI(t) [+ γ·EarningsDummy]",
        "variable_definitions": {
            "us_lag1": [{"key": v["key"], "name": v["name"], "ticker": v["ticker"]} for v in US_VARIABLES],
            "kr_lag0": [{"key": v["key"], "name": v["name"], "ticker": v["ticker"]} for v in KR_VARIABLES],
            "event_dummy": "어닝 발표일 ±2 거래일에 1, 평소 0",
        },
        "kr_stocks": [{"ticker": k["ticker"], "name": k["name"]} for k in KR_SEMIS],
        "results": results,
        "charts": {
            "normalized_prices": "charts/semi_normalized.png",
            "cumulative_returns": "charts/semi_cumulative.png",
        },
        "warning": "🚨 다변량 회귀도 통계적 추정에 불과. 실거래 결정에 단독 사용 금지.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    # Telegram 요약
    if results:
        lines = ["📊 반도체 민감도 v2 (다변량)", "=" * 30, ""]
        for r in results[:6]:
            mv = r.get("multivariate") or {}
            if not mv:
                continue
            b = mv.get("betas", {})
            mv_ev = r.get("multivariate_with_event") or {}
            ev_tag = ""
            if mv_ev:
                d = mv_ev.get("r_squared", 0) - mv.get("r_squared", 0)
                ev_tag = f" (+이벤트 ΔR² {d:+.2f})"
            lines.append(f"{r['name']} R²={mv.get('r_squared', 0):.2f}{ev_tag}")
            lines.append(f"  NQ {b.get('nq', 0):+.2f} | SOX {b.get('sox', 0):+.2f} | KOSPI {b.get('kospi', 0):+.2f}")
            lines.append(f"  USDKRW {b.get('usdkrw', 0):+.2f} | UST10Y {b.get('ust10y', 0):+.2f}")
            lines.append("")
        lines.append("🚨 통계 추정. 자동 매매 금지.")
        lines.append("\n대시보드: https://15678910.github.io/ai-finance/")
        send_message("\n".join(lines))

    print("\n" + "=" * 72)
    print(f"  v2 분석 완료: {len(results)}개 종목")
    print("=" * 72)


if __name__ == "__main__":
    main()
