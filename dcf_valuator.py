"""
DCF (현금흐름할인) 자동 평가기
================================

Anthropic Claude for Financial Services의 'Model Builder Agent'에서 영감을 받아
종목별 적정 주가를 DCF 모델로 자동 산출합니다.

평가 절차:
  1. 과거 5년 매출/잉여현금흐름 추출 (yfinance)
  2. CAGR 계산하여 향후 5년 FCF 추정
  3. WACC 계산 (CAPM 기반)
  4. 잔여가치(Terminal Value) — 영구성장법(기본) + EV/EBITDA 출구배수법(교차검증)
  5. 연중 할인(mid-year convention)으로 현재가치 → 적정 시총
  6. 희석주식수로 나눠서 적정 주가
  7. 현재가 대비 ±% 비교, 두 터미널 방법의 괴리 표시
  8. 민감도 분석 (할인율, 성장률 ±)

투자은행 실무 관행 반영 (2026-09):
  · 연중 할인 — 연말 할인은 모든 흐름을 반년 늦게 받는 셈이라 WACC 9%에서 약 4.4% 저평가
  · 터미널 이원화 — EV의 70%+가 영구성장률 가정 하나에 매달리므로 시장 배수로 교차검증
  · 희석주식수 — 기본주식수로 나누면 스톡옵션·RSU 희석분만큼 적정주가가 과대 산출

알려진 한계: 리스부채를 순부채에 넣지 않는다. 현재 대상(제조·바이오·IT)에는 영향이
작지만 유통·항공(이마트·대한항공 등)을 추가하면 반드시 함께 처리해야 한다.

🚨 시뮬레이션 전용. 단순화된 가정 사용. 실제 매매 금지.
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
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "dcf_valuations.json")
KST = timezone(timedelta(hours=9))

# DCF 가정값
RISK_FREE_RATE = 0.035        # 한국 10Y 국채 3.5%
EQUITY_RISK_PREMIUM = 0.06    # 주식 위험 프리미엄 6%
TERMINAL_GROWTH = 0.02         # 영구 성장률 2%
PROJECTION_YEARS = 5           # 명시적 추정 기간
DEFAULT_TAX_RATE = 0.22        # 한국 법인세율 22%

# 연중 할인 (mid-year convention)
# 현금흐름은 연말에 한꺼번에 들어오지 않고 한 해에 걸쳐 발생한다. 연말 할인은
# 모든 흐름을 반년씩 늦게 받는 셈이라 가치를 체계적으로 낮춘다 — WACC 9%에서
# 약 -4.4%. 투자은행 실무 표준이며, 이전 코드는 이 보정이 없었다.
MID_YEAR_CONVENTION = True

# 터미널 가치 출구배수법 (EV/EBITDA exit multiple)
# 영구성장법 하나로 터미널 가치를 구하면 EV의 70~80%가 '영구성장률 2%' 가정
# 하나에 매달린다. 시장 배수로 다시 계산해 두 방법의 괴리를 드러낸다.
DEFAULT_EXIT_MULTIPLE = 8.0            # 시장 배수 없을 때 폴백 (KOSPI 대형주 평균권)
EXIT_MULTIPLE_RANGE = (4.0, 20.0)      # 이 범위 밖 시장 배수는 이상치로 보고 폴백

# 평가 대상 (KOSPI 대형주)
TARGET_STOCKS = [
    {"name": "삼성전자", "ticker": "005930.KS"},
    {"name": "SK하이닉스", "ticker": "000660.KS"},
    {"name": "한미반도체", "ticker": "042700.KS"},
    {"name": "LG에너지솔루션", "ticker": "373220.KS"},
    {"name": "삼성SDI", "ticker": "006400.KS"},
    {"name": "현대차", "ticker": "005380.KS"},
    {"name": "기아", "ticker": "000270.KS"},
    {"name": "POSCO홀딩스", "ticker": "005490.KS"},
    {"name": "한화에어로스페이스", "ticker": "012450.KS"},
    {"name": "LIG넥스원", "ticker": "079550.KS"},
    {"name": "셀트리온", "ticker": "068270.KS"},
    {"name": "삼성바이오로직스", "ticker": "207940.KS"},
    {"name": "네이버", "ticker": "035420.KS"},
    {"name": "현대로템", "ticker": "064350.KS"},
    {"name": "한미약품", "ticker": "128940.KS"},
]


# ====================================================================
# DCF 핵심 함수
# ====================================================================
def calculate_wacc(beta: float, tax_rate: float = DEFAULT_TAX_RATE,
                   debt_ratio: float = 0.3) -> float:
    """가중평균자본비용 (WACC) 계산.
    단순화: 부채/자본 = 30%/70% 가정.
    """
    if beta is None or beta <= 0:
        beta = 1.0
    cost_of_equity = RISK_FREE_RATE + beta * EQUITY_RISK_PREMIUM
    cost_of_debt = RISK_FREE_RATE + 0.02  # 회사채 스프레드 +2%p
    after_tax_debt = cost_of_debt * (1 - tax_rate)
    wacc = (1 - debt_ratio) * cost_of_equity + debt_ratio * after_tax_debt
    return wacc


def growth_path(growth_rate: float, years: int = PROJECTION_YEARS) -> list:
    """연도별 성장률 경로 — 초기 성장률에서 영구성장률로 선형 수렴.

    FCF와 EBITDA 추정이 같은 경로를 써야 영구성장법과 출구배수법의 터미널 가치를
    같은 잣대로 비교할 수 있다.
    """
    path = []
    for year in range(1, years + 1):
        decay = (growth_rate - TERMINAL_GROWTH) * (1 - year / years)
        path.append(TERMINAL_GROWTH + decay)
    return path


def project_fcf(historical_fcf: list, growth_rate: float, years: int = PROJECTION_YEARS) -> list:
    """과거 FCF + 추정 성장률로 미래 FCF 추정."""
    if not historical_fcf:
        return []
    current = historical_fcf[-1]  # 가장 최근 FCF
    projected = []
    for g in growth_path(growth_rate, years):
        current = current * (1 + g)
        projected.append(current)
    return projected


def calculate_dcf(projected_fcf: list, wacc: float, terminal_growth: float = TERMINAL_GROWTH,
                  mid_year: bool = MID_YEAR_CONVENTION,
                  terminal_ebitda: float = None, exit_multiple: float = None) -> dict:
    """DCF 평가 → 현재가치 (Enterprise Value).

    할인 시점 규칙 (mid_year=True):
      · 명시 기간 FCF   → t - 0.5   (연중 발생)
      · 영구성장 터미널 → N - 0.5   (N+1년 흐름부터의 영구 스트림을 같은 관행으로 할인)
      · 출구배수 터미널 → N         (연말에 EBITDA×배수로 매각한다는 가정이라 연말 시점)
    출구배수 터미널을 연중으로 당기면 매각 시점을 반년 앞당기는 셈이라 과대평가된다.
    """
    empty = {"enterprise_value": 0, "pv_of_fcf": 0, "pv_of_terminal": 0, "terminal_value": 0,
             "terminal_share_pct": None, "terminal_value_exit": None,
             "pv_of_terminal_exit": None, "enterprise_value_exit": None}
    if not projected_fcf or wacc <= terminal_growth:
        return empty

    n = len(projected_fcf)
    offset = 0.5 if mid_year else 0.0

    # 명시적 기간 FCF 현재가치
    pv_of_fcf = sum(fcf / ((1 + wacc) ** (i + 1 - offset)) for i, fcf in enumerate(projected_fcf))

    # 잔여가치 ① 영구성장법 (Gordon growth) — 기본값이자 대시보드 표시값
    terminal_value = projected_fcf[-1] * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_of_terminal = terminal_value / ((1 + wacc) ** (n - offset))
    enterprise_value = pv_of_fcf + pv_of_terminal

    result = {
        "enterprise_value": enterprise_value,
        "pv_of_fcf": pv_of_fcf,
        "pv_of_terminal": pv_of_terminal,
        "terminal_value": terminal_value,
        # EV 중 터미널 비중 — 클수록 영구성장률 가정 하나에 매달린 평가
        "terminal_share_pct": (pv_of_terminal / enterprise_value * 100) if enterprise_value else None,
        "terminal_value_exit": None,
        "pv_of_terminal_exit": None,
        "enterprise_value_exit": None,
    }

    # 잔여가치 ② 출구배수법 (EV/EBITDA exit multiple) — 교차검증용
    if terminal_ebitda and exit_multiple and terminal_ebitda > 0 and exit_multiple > 0:
        tv_exit = terminal_ebitda * exit_multiple
        pv_exit = tv_exit / ((1 + wacc) ** n)       # 연말 매각 가정 → 오프셋 없음
        result.update({
            "terminal_value_exit": tv_exit,
            "pv_of_terminal_exit": pv_exit,
            "enterprise_value_exit": pv_of_fcf + pv_exit,
        })

    return result


def select_share_count(fin: dict) -> tuple:
    """(주식수, 기준). 희석주식수를 우선한다.

    sharesOutstanding은 기본주식수라 스톡옵션·RSU 희석분이 빠진다. 분모가 작으면
    적정주가가 과대 산출된다. impliedSharesOutstanding(=시총/주가)이 Yahoo가 실제로
    시총 산출에 쓰는 주식수라 더 완전하다. 두 값이 2배 이상 어긋나면 단위 오류로
    보고 기본주식수로 돌아간다.
    """
    implied = float(fin.get("shares_implied") or 0)
    basic = float(fin.get("shares_outstanding") or 0)
    if implied > 0 and basic > 0:
        ratio = implied / basic
        if 0.5 <= ratio <= 2.0:
            return implied, "implied_diluted"
        return basic, "basic (implied 이상치)"
    if implied > 0:
        return implied, "implied_diluted"
    if basic > 0:
        return basic, "basic"
    return 0.0, "none"


def select_exit_multiple(fin: dict) -> tuple:
    """(EV/EBITDA 출구배수, 출처). 현재 시장 배수를 쓰되 범위 밖이면 기본값.

    시장 배수를 쓰면 현재 밸류에이션이 터미널에 그대로 박히는 순환성이 있다.
    그래도 '시장이 이 회사에 매기는 배수'는 영구성장률 2%보다 관측 가능한 가정이라
    교차검증 축으로는 낫다. 출처를 함께 남겨 어느 쪽인지 알 수 있게 한다.
    """
    lo, hi = EXIT_MULTIPLE_RANGE
    m = fin.get("ev_to_ebitda")
    try:
        m = float(m) if m is not None else None
    except (TypeError, ValueError):
        m = None
    if m is not None and lo <= m <= hi:
        return m, "market"
    return DEFAULT_EXIT_MULTIPLE, "default"


def calculate_cagr(values: list, years: int) -> float:
    """연복합성장률 (CAGR)."""
    if not values or len(values) < 2 or years <= 0:
        return 0.0
    start = values[0]
    end = values[-1]
    if start <= 0:
        return 0.0
    cagr = (end / start) ** (1 / years) - 1
    return cagr


# ====================================================================
# yfinance에서 재무 데이터 추출
# ====================================================================
def fetch_financials(ticker: str) -> dict:
    """yfinance에서 재무 데이터 추출."""
    try:
        t = yf.Ticker(ticker)
        info = t.info or {}
        cashflow = t.cashflow  # 현금흐름표

        result = {
            "current_price": info.get("currentPrice") or info.get("regularMarketPrice"),
            "market_cap": info.get("marketCap"),
            "shares_outstanding": info.get("sharesOutstanding"),      # 기본주식수
            "shares_implied": info.get("impliedSharesOutstanding"),   # 희석 반영 (시총/주가)
            "beta": info.get("beta") or 1.0,
            "total_debt": info.get("totalDebt", 0),
            "cash": info.get("totalCash", 0),
            "trailing_pe": info.get("trailingPE"),
            "forward_pe": info.get("forwardPE"),
            "revenue": info.get("totalRevenue"),
            # 출구배수법용 — 최근 EBITDA와 시장이 매기는 EV/EBITDA
            "ebitda": info.get("ebitda"),
            "ev_to_ebitda": info.get("enterpriseToEbitda"),
        }

        # 잉여현금흐름 (FCF) = 영업현금흐름 - 자본적지출
        fcf_history = []
        if cashflow is not None and not cashflow.empty:
            for col in cashflow.columns[:5]:  # 최근 5년
                ocf_keys = ['Operating Cash Flow', 'Total Cash From Operating Activities']
                capex_keys = ['Capital Expenditure', 'Capital Expenditures']

                ocf = None
                for k in ocf_keys:
                    if k in cashflow.index:
                        ocf = cashflow.loc[k, col]
                        break
                capex = None
                for k in capex_keys:
                    if k in cashflow.index:
                        capex = cashflow.loc[k, col]
                        break

                if ocf is not None and pd.notna(ocf):
                    capex_val = float(capex) if capex is not None and pd.notna(capex) else 0
                    fcf = float(ocf) + capex_val  # capex는 보통 음수
                    fcf_history.append(fcf)

        result["fcf_history"] = list(reversed(fcf_history))  # 오래된 → 최근 순

        # 매출 히스토리
        income_stmt = t.income_stmt
        revenue_history = []
        if income_stmt is not None and not income_stmt.empty:
            for col in income_stmt.columns[:5]:
                rev_keys = ['Total Revenue', 'Revenue']
                for k in rev_keys:
                    if k in income_stmt.index:
                        v = income_stmt.loc[k, col]
                        if pd.notna(v):
                            revenue_history.append(float(v))
                        break
        result["revenue_history"] = list(reversed(revenue_history))

        return result
    except Exception as e:
        print(f"  [실패] {ticker}: {e}")
        return {}


# ====================================================================
# 평가 실행
# ====================================================================
def evaluate_stock(name: str, ticker: str) -> dict:
    """단일 종목 DCF 평가."""
    print(f"\n  [{name}] 평가 중...")

    fin = fetch_financials(ticker)
    if not fin or not fin.get("current_price"):
        print(f"    [건너뜀] 데이터 부족")
        return None

    fcf_history = fin.get("fcf_history", [])
    revenue_history = fin.get("revenue_history", [])

    if len(fcf_history) < 2:
        print(f"    [건너뜀] FCF 데이터 부족 ({len(fcf_history)}년)")
        return None

    # 매출 CAGR (성장률 추정)
    if len(revenue_history) >= 2:
        rev_cagr = calculate_cagr(revenue_history, len(revenue_history) - 1)
    else:
        rev_cagr = 0.05  # 기본 5%

    # 합리적 범위로 제한
    growth_rate = max(0.02, min(0.30, rev_cagr))

    # WACC 계산
    beta = fin.get("beta", 1.0)
    wacc = calculate_wacc(beta)

    # 미래 FCF 추정
    projected_fcf = project_fcf(fcf_history, growth_rate)

    # 출구배수법 입력 — 최근 EBITDA를 FCF와 같은 성장 경로로 N년차까지 키운다
    exit_multiple, exit_source = select_exit_multiple(fin)
    ebitda_latest = float(fin.get("ebitda") or 0)
    terminal_ebitda = None
    if ebitda_latest > 0:
        cum = 1.0
        for g in growth_path(growth_rate):
            cum *= (1 + g)
        terminal_ebitda = ebitda_latest * cum

    # DCF 평가 (연중 할인 + 영구성장·출구배수 이원화)
    dcf = calculate_dcf(projected_fcf, wacc,
                        terminal_ebitda=terminal_ebitda, exit_multiple=exit_multiple)
    enterprise_value = dcf["enterprise_value"]

    # Equity Value = EV - 순부채
    net_debt = (fin.get("total_debt") or 0) - (fin.get("cash") or 0)
    equity_value = enterprise_value - net_debt

    # 적정 주가 — 희석주식수 우선
    shares, shares_basis = select_share_count(fin)
    if shares <= 0:
        return None
    fair_price = equity_value / shares

    # 출구배수법 적정주가 — 교차검증. 두 방법의 괴리가 크면 터미널 가정이 취약한 것
    fair_price_exit = None
    terminal_method_gap_pct = None
    if dcf.get("enterprise_value_exit") is not None:
        fair_price_exit = (dcf["enterprise_value_exit"] - net_debt) / shares
        if fair_price > 0:
            terminal_method_gap_pct = (fair_price_exit - fair_price) / fair_price * 100
    fair_price_blend = ((fair_price + fair_price_exit) / 2
                        if fair_price_exit is not None else fair_price)

    # 현재가 대비 상하방
    current_price = fin.get("current_price", 0)
    upside_pct = (fair_price - current_price) / current_price * 100 if current_price > 0 else 0

    # 민감도 분석 (WACC ±1%, 성장률 ±2%)
    sensitivity = {}
    for d_wacc in [-0.01, 0, 0.01]:
        for d_growth in [-0.02, 0, 0.02]:
            adj_wacc = wacc + d_wacc
            adj_growth = max(0.02, growth_rate + d_growth)
            adj_fcf = project_fcf(fcf_history, adj_growth)
            adj_dcf = calculate_dcf(adj_fcf, adj_wacc)
            adj_equity = adj_dcf["enterprise_value"] - net_debt
            adj_fair = adj_equity / shares if shares > 0 else 0
            adj_upside = (adj_fair - current_price) / current_price * 100 if current_price > 0 else 0
            sensitivity[f"wacc{d_wacc:+.2f}_g{d_growth:+.2f}"] = {
                "fair_price": round(adj_fair, 0),
                "upside_pct": round(adj_upside, 1),
            }

    # 평가 시그널
    if upside_pct > 30:
        signal = "🟢 강한 매수 시그널"
    elif upside_pct > 15:
        signal = "🟢 저평가"
    elif upside_pct > -15:
        signal = "🟡 적정 수준"
    elif upside_pct > -30:
        signal = "🟠 고평가"
    else:
        signal = "🔴 심각 고평가"

    return {
        "name": name,
        "ticker": ticker,
        "current_price": round(current_price, 0),
        "fair_price_dcf": round(fair_price, 0),
        "upside_pct": round(upside_pct, 1),
        "signal": signal,
        "wacc_pct": round(wacc * 100, 2),
        "growth_rate_pct": round(growth_rate * 100, 2),
        "beta": round(beta, 2),
        "fcf_latest": fcf_history[-1] if fcf_history else 0,
        "fcf_history": fcf_history,
        "revenue_cagr_pct": round(rev_cagr * 100, 2) if len(revenue_history) >= 2 else None,
        "enterprise_value": round(enterprise_value, 0),
        "equity_value": round(equity_value, 0),
        "market_cap": fin.get("market_cap", 0),
        "trailing_pe": fin.get("trailing_pe"),
        "forward_pe": fin.get("forward_pe"),
        "shares_outstanding": shares,
        "shares_basis": shares_basis,
        # 터미널 가치 이원화 — fair_price_dcf(영구성장법)는 그대로 두고 병기한다
        "terminal_share_pct": (round(dcf["terminal_share_pct"], 1)
                               if dcf.get("terminal_share_pct") is not None else None),
        "fair_price_exit": round(fair_price_exit, 0) if fair_price_exit is not None else None,
        "fair_price_blend": round(fair_price_blend, 0),
        "terminal_method_gap_pct": (round(terminal_method_gap_pct, 1)
                                    if terminal_method_gap_pct is not None else None),
        "exit_multiple": round(exit_multiple, 2),
        "exit_multiple_source": exit_source,
        "ebitda_latest": ebitda_latest or None,
        "mid_year_convention": MID_YEAR_CONVENTION,
        "sensitivity": sensitivity,
    }


# ====================================================================
# 텔레그램 (저평가 Top 5만)
# ====================================================================
def send_telegram_summary(valuations: list):
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

    # 강한 매수 시그널만 전송
    strong_buys = [v for v in valuations if v and v.get("upside_pct", 0) > 30][:5]
    if not strong_buys:
        return

    lines = ["💎 DCF 저평가 Top 5", "=" * 25, ""]
    for i, v in enumerate(strong_buys, 1):
        lines.append(f"{i}. {v['name']} ({v['ticker'].split('.')[0]})")
        lines.append(f"   현재가: {v['current_price']:,}원")
        lines.append(f"   적정가: {v['fair_price_dcf']:,}원")
        lines.append(f"   상방: {v['upside_pct']:+.1f}% {v['signal']}")
        lines.append(f"   WACC {v['wacc_pct']:.1f}% / 성장률 {v['growth_rate_pct']:.1f}%")
        lines.append("")

    lines.append("🚨 단순화된 DCF 가정 사용. 자동 매매 금지.")
    lines.append("\n대시보드: https://15678910.github.io/ai-finance/")

    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        body = urllib.parse.urlencode({"chat_id": chat_id, "text": "\n".join(lines)}).encode()
        req = urllib.request.Request(url, data=body, method="POST")
        with urllib.request.urlopen(req, timeout=10) as resp:
            json.loads(resp.read())
        print(f"  [텔레그램] DCF 저평가 {len(strong_buys)}건 전송")
    except Exception as e:
        print(f"  [텔레그램] 실패: {e}")


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 65)
    print("  DCF 자동 평가기 (Discounted Cash Flow Valuation)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)
    print(f"  무위험금리: {RISK_FREE_RATE*100:.1f}%, 시장프리미엄: {EQUITY_RISK_PREMIUM*100:.1f}%")
    print(f"  영구성장률: {TERMINAL_GROWTH*100:.1f}%, 추정기간: {PROJECTION_YEARS}년")
    print("=" * 65)

    valuations = []
    for stock in TARGET_STOCKS:
        try:
            v = evaluate_stock(stock["name"], stock["ticker"])
            if v:
                valuations.append(v)
                print(f"    적정가 {v['fair_price_dcf']:,.0f}원 / 현재가 {v['current_price']:,.0f}원 → {v['upside_pct']:+.1f}% {v['signal']}")
        except Exception as e:
            print(f"    [오류] {stock['name']}: {e}")

    # 정렬: upside_pct 내림차순
    valuations.sort(key=lambda x: x.get("upside_pct", 0), reverse=True)

    # 결과 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "assumptions": {
            "risk_free_rate_pct": RISK_FREE_RATE * 100,
            "equity_risk_premium_pct": EQUITY_RISK_PREMIUM * 100,
            "terminal_growth_pct": TERMINAL_GROWTH * 100,
            "projection_years": PROJECTION_YEARS,
            "tax_rate_pct": DEFAULT_TAX_RATE * 100,
            "mid_year_convention": MID_YEAR_CONVENTION,
            "exit_multiple_default": DEFAULT_EXIT_MULTIPLE,
            "exit_multiple_range": list(EXIT_MULTIPLE_RANGE),
            "shares_basis": "희석주식수(impliedSharesOutstanding) 우선, 없으면 기본주식수",
        },
        "valuations": valuations,
        "warning": "🚨 단순화된 DCF 가정. 시뮬레이션 전용. 자동 매매 금지.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    # 텔레그램
    send_telegram_summary(valuations)

    print("\n" + "=" * 65)
    print(f"  평가 완료: {len(valuations)}개 종목")
    if valuations:
        top_under = valuations[0]
        print(f"  최고 저평가: {top_under['name']} ({top_under['upside_pct']:+.1f}%)")
    print("  ⚠️ 시뮬레이션 / 단순화된 가정. 자동 매매 금지.")
    print("=" * 65)


if __name__ == "__main__":
    main()
