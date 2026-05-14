"""
한국 주식 IRR (Internal Rate of Return) 분석기
==================================================

목적: 종목별 기대 연수익률을 다각도로 추정 + 위험 조정 + 섹터별 비교.

Phase C (value_screener.py)의 기본 Gordon IRR 위에 다음을 추가:
  1. 다중 IRR 방법 통합 (Gordon, Earnings Yield, ROIC-Retention)
  2. 위험 조정 (베타, 부채비율, 매출 변동성, 코리아 디스카운트)
  3. 섹터별 IRR 분포 (평균·중앙값·표준편차)
  4. 안전 마진 분석 (IRR - 필요수익률)
  5. 텔레그램 자동 알림: 강력 매수 후보 (IRR > 12% + 스프레드 > 7%pt + ROE > 8%)

방법론:
  필요수익률(Required Return) = 무위험금리 + β × 시장프리미엄 + 코리아 디스카운트
    · 무위험금리: 한국 국고채 10년 ~3.5%
    · 시장프리미엄: 6.0% (글로벌 평균)
    · 코리아 디스카운트: +1.5%pt (지배구조·소액주주권리 약점)

  안전마진(MoS) = Primary IRR - Required Return
    · MoS > 5%pt: 강력 매수 후보
    · MoS 2-5%pt: 매수 검토
    · MoS 0-2%pt: 시장 평균
    · MoS < 0: 고평가

🚨 시뮬레이션 / 통계 추정. 실거래 단독 사용 금지.
"""

import os
import sys
import json
from datetime import datetime, timezone, timedelta
from collections import defaultdict

try:
    import yfinance as yf
    import numpy as np
except ImportError as e:
    print(f"[오류] 라이브러리 미설치: {e}")
    sys.exit(1)

from core import send_message, load_state, save_state

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "irr_analysis.json")
STATE_NAME = "irr_analysis"
KST = timezone(timedelta(hours=9))

# ====================================================================
# 시장 가정 (한국 시장 기준)
# ====================================================================
RISK_FREE_KR = 0.035        # 한국 국고채 10년
EQUITY_PREMIUM = 0.060      # 시장 위험 프리미엄 (글로벌 평균)
KOREA_DISCOUNT = 0.015      # 코리아 디스카운트 (지배구조·소액주주권리 약점)
GROWTH_CAP_HIGH = 0.10      # 단기 호황 종목 영구성장 상단
GROWTH_CAP_LOW = -0.05      # 침체 종목 영구성장 하단
GROWTH_DEFAULT = 0.03       # 데이터 부족 시 디폴트

# 알림 트리거 조건 (강력 매수 후보)
STRONG_BUY_IRR = 0.12       # IRR 12% 이상
STRONG_BUY_MOS = 0.05       # 안전마진 5%pt 이상
STRONG_BUY_ROE = 0.08       # ROE 8% 이상 (이익품질)

# ====================================================================
# 분석 대상 종목 (value_screener의 universe 재사용 + 확장)
# ====================================================================
TARGET_STOCKS = [
    # IT/반도체
    {"ticker": "005930.KS", "name": "삼성전자", "sector": "IT/반도체"},
    {"ticker": "000660.KS", "name": "SK하이닉스", "sector": "IT/반도체"},
    {"ticker": "042700.KS", "name": "한미반도체", "sector": "반도체장비"},
    {"ticker": "000990.KS", "name": "DB하이텍", "sector": "반도체"},
    {"ticker": "035420.KS", "name": "네이버", "sector": "IT/플랫폼"},
    {"ticker": "035720.KS", "name": "카카오", "sector": "IT/플랫폼"},
    {"ticker": "066570.KS", "name": "LG전자", "sector": "전자"},
    {"ticker": "018260.KS", "name": "삼성에스디에스", "sector": "IT"},
    {"ticker": "034020.KS", "name": "두산에너빌리티", "sector": "IT/원전"},

    # 금융
    {"ticker": "055550.KS", "name": "신한지주", "sector": "금융"},
    {"ticker": "086790.KS", "name": "하나금융지주", "sector": "금융"},
    {"ticker": "105560.KS", "name": "KB금융", "sector": "금융"},
    {"ticker": "316140.KS", "name": "우리금융지주", "sector": "금융"},
    {"ticker": "024110.KS", "name": "기업은행", "sector": "금융"},
    {"ticker": "138930.KS", "name": "BNK금융지주", "sector": "금융"},
    {"ticker": "071050.KS", "name": "한국금융지주", "sector": "증권"},
    {"ticker": "006800.KS", "name": "미래에셋증권", "sector": "증권"},

    # 자동차
    {"ticker": "005380.KS", "name": "현대차", "sector": "자동차"},
    {"ticker": "000270.KS", "name": "기아", "sector": "자동차"},
    {"ticker": "012330.KS", "name": "현대모비스", "sector": "자동차"},
    {"ticker": "204320.KS", "name": "HL만도", "sector": "자동차부품"},

    # 화학/에너지
    {"ticker": "051910.KS", "name": "LG화학", "sector": "화학"},
    {"ticker": "010950.KS", "name": "S-Oil", "sector": "에너지"},
    {"ticker": "096770.KS", "name": "SK이노베이션", "sector": "에너지"},
    {"ticker": "015760.KS", "name": "한국전력", "sector": "유틸리티"},
    {"ticker": "036460.KS", "name": "한국가스공사", "sector": "유틸리티"},

    # 철강/조선/방산
    {"ticker": "005490.KS", "name": "POSCO홀딩스", "sector": "철강"},
    {"ticker": "009540.KS", "name": "HD한국조선해양", "sector": "조선"},
    {"ticker": "010140.KS", "name": "삼성중공업", "sector": "조선"},
    {"ticker": "042660.KS", "name": "한화오션", "sector": "조선"},
    {"ticker": "012450.KS", "name": "한화에어로스페이스", "sector": "방산"},
    {"ticker": "079550.KS", "name": "LIG넥스원", "sector": "방산"},
    {"ticker": "064350.KS", "name": "현대로템", "sector": "방산"},

    # 바이오/제약
    {"ticker": "068270.KS", "name": "셀트리온", "sector": "바이오"},
    {"ticker": "207940.KS", "name": "삼성바이오로직스", "sector": "바이오"},
    {"ticker": "128940.KS", "name": "한미약품", "sector": "제약"},
    {"ticker": "000100.KS", "name": "유한양행", "sector": "제약"},
    {"ticker": "009420.KS", "name": "한올바이오파마", "sector": "바이오"},

    # 배터리/2차전지
    {"ticker": "373220.KS", "name": "LG에너지솔루션", "sector": "배터리"},
    {"ticker": "006400.KS", "name": "삼성SDI", "sector": "배터리"},
    {"ticker": "247540.KQ", "name": "에코프로비엠", "sector": "배터리소재"},

    # 통신/유틸
    {"ticker": "017670.KS", "name": "SK텔레콤", "sector": "통신"},
    {"ticker": "030200.KS", "name": "KT", "sector": "통신"},
    {"ticker": "032640.KS", "name": "LG유플러스", "sector": "통신"},

    # 유통/소비재
    {"ticker": "271560.KS", "name": "오리온", "sector": "소비재"},
    {"ticker": "097950.KS", "name": "CJ제일제당", "sector": "소비재"},
    {"ticker": "139480.KS", "name": "이마트", "sector": "유통"},
    {"ticker": "023530.KS", "name": "롯데쇼핑", "sector": "유통"},
    {"ticker": "008770.KS", "name": "호텔신라", "sector": "관광"},
    {"ticker": "033780.KS", "name": "KT&G", "sector": "고배당"},

    # 건설/물류
    {"ticker": "000720.KS", "name": "현대건설", "sector": "건설"},
    {"ticker": "047040.KS", "name": "대우건설", "sector": "건설"},
    {"ticker": "000120.KS", "name": "CJ대한통운", "sector": "물류"},

    # 게임/엔터
    {"ticker": "036570.KS", "name": "엔씨소프트", "sector": "게임"},
    {"ticker": "259960.KS", "name": "크래프톤", "sector": "게임"},
    {"ticker": "352820.KS", "name": "하이브", "sector": "엔터"},
    {"ticker": "035900.KQ", "name": "JYP엔터테인먼트", "sector": "엔터"},
    {"ticker": "041510.KQ", "name": "에스엠", "sector": "엔터"},

    # 항공
    {"ticker": "003490.KS", "name": "대한항공", "sector": "항공"},

    # 화장품
    {"ticker": "090430.KS", "name": "아모레퍼시픽", "sector": "화장품"},
    {"ticker": "051900.KS", "name": "LG생활건강", "sector": "화장품"},
]


# ====================================================================
# yfinance 데이터 형식 헬퍼 (Phase C와 동일 로직)
# ====================================================================
def _norm_div_yield(raw):
    """yfinance dividendYield(% 포맷) → 소수."""
    if raw is None:
        return None
    return raw / 100.0


def _cap_growth(raw, lo=-0.30, hi=0.50):
    """비정상 성장률 필터 (저기저 효과 제외)."""
    if raw is None:
        return None
    if raw > hi or raw < lo:
        return None
    return raw


def fetch_data(ticker: str) -> dict:
    """단일 종목 데이터 수집."""
    try:
        t = yf.Ticker(ticker)
        info = t.info or {}
        return {
            "current_price": info.get("currentPrice") or info.get("regularMarketPrice"),
            "market_cap": info.get("marketCap"),
            "trailing_pe": info.get("trailingPE"),
            "forward_pe": info.get("forwardPE"),
            "price_to_book": info.get("priceToBook"),
            "roe": info.get("returnOnEquity"),  # 소수
            "operating_margin": info.get("operatingMargins"),
            "debt_to_equity": info.get("debtToEquity"),  # % 단위
            "dividend_yield_raw": info.get("dividendYield"),  # % 단위
            "revenue_growth": info.get("revenueGrowth"),  # 소수
            "earnings_growth": info.get("earningsGrowth"),  # 소수
            "beta": info.get("beta"),
            "fcf": info.get("freeCashflow"),
            "shares": info.get("sharesOutstanding"),
        }
    except Exception as e:
        print(f"  [실패] {ticker}: {e}")
        return {}


# ====================================================================
# IRR 계산 (3가지 방법)
# ====================================================================
def estimate_growth(data: dict) -> tuple:
    """영구성장률(g) 추정 + 출처 라벨."""
    rg = _cap_growth(data.get("revenue_growth"))
    eg = _cap_growth(data.get("earnings_growth"))

    if rg is not None and eg is not None:
        # 둘 중 보수적(낮은 값) 선택 (장기 영구 가정)
        g = min(rg, eg)
        src = "min(매출, 이익) 성장률"
    elif rg is not None:
        g = rg
        src = "매출 성장률"
    elif eg is not None:
        g = eg
        src = "이익 성장률"
    else:
        # 지속가능성장률 모델
        roe = data.get("roe")
        div_yield = _norm_div_yield(data.get("dividend_yield_raw"))
        pe = data.get("trailing_pe") or data.get("forward_pe")
        if roe is not None and div_yield is not None and pe and pe > 0:
            payout = min(max(div_yield * pe, 0), 1.0)
            g = roe * (1 - payout)
            src = f"지속가능성장률 (ROE×retention)"
        else:
            g = GROWTH_DEFAULT
            src = "디폴트 3%"

    g = max(min(g, GROWTH_CAP_HIGH), GROWTH_CAP_LOW)
    return g, src


def calculate_irrs(data: dict) -> dict:
    """3가지 IRR 동시 계산."""
    div_yield = _norm_div_yield(data.get("dividend_yield_raw"))
    pe = data.get("trailing_pe") or data.get("forward_pe")
    roe = data.get("roe")

    g, g_src = estimate_growth(data)

    # 1) Gordon Growth IRR
    gordon = None
    if div_yield and div_yield > 0.001:
        gordon = div_yield + g

    # 2) Earnings Yield
    earnings_yield = None
    if pe and pe > 0:
        earnings_yield = 1 / pe

    # 3) ROIC-Retention IRR (ROE × retention + 배당수익률)
    roic_retention = None
    if roe is not None and div_yield is not None and pe and pe > 0:
        payout = min(max(div_yield * pe, 0), 1.0)
        sustainable_g = roe * (1 - payout)
        sustainable_g = min(max(sustainable_g, GROWTH_CAP_LOW), GROWTH_CAP_HIGH)
        roic_retention = div_yield + sustainable_g

    # Primary IRR 결정: 안전하게 보수적 평균
    candidates = [v for v in [gordon, earnings_yield, roic_retention] if v is not None]
    if candidates:
        # 평균 + 중앙값 둘 다 계산
        primary_irr = float(np.median(candidates))
        method = "median of available methods"
    else:
        primary_irr = None
        method = None

    return {
        "gordon_pct": round(gordon * 100, 2) if gordon is not None else None,
        "earnings_yield_pct": round(earnings_yield * 100, 2) if earnings_yield is not None else None,
        "roic_retention_pct": round(roic_retention * 100, 2) if roic_retention is not None else None,
        "primary_irr_pct": round(primary_irr * 100, 2) if primary_irr is not None else None,
        "implied_growth_pct": round(g * 100, 2),
        "growth_source": g_src,
        "method": method,
    }


# ====================================================================
# 위험 조정 — 필요수익률(Required Return) 계산
# ====================================================================
def calculate_required_return(data: dict, sector: str) -> dict:
    """CAPM + 한국 시장 보정."""
    beta = data.get("beta")
    debt_to_equity = data.get("debt_to_equity")  # % 단위

    # 베타: 데이터 없으면 1.0 가정 (시장 평균)
    if beta is None or beta < 0:
        beta = 1.0
        beta_source = "default 1.0"
    else:
        beta_source = f"yfinance ({beta:.2f})"

    # 기본 CAPM
    base_required = RISK_FREE_KR + beta * EQUITY_PREMIUM

    # 한국 디스카운트
    korea_adjustment = KOREA_DISCOUNT

    # 부채비율 페널티 (D/E > 200% 시 추가 위험)
    debt_penalty = 0.0
    if debt_to_equity is not None:
        if debt_to_equity > 200:
            debt_penalty = 0.01
        elif debt_to_equity > 150:
            debt_penalty = 0.005

    # 섹터별 추가 위험 (cyclical/배터리 등 변동성 큰 섹터)
    sector_penalty = 0.0
    high_vol_sectors = {"배터리", "배터리소재", "조선", "철강", "에너지", "건설", "엔터", "게임", "바이오"}
    if sector in high_vol_sectors:
        sector_penalty = 0.005

    total = base_required + korea_adjustment + debt_penalty + sector_penalty

    return {
        "required_return_pct": round(total * 100, 2),
        "base_capm_pct": round(base_required * 100, 2),
        "korea_discount_pct": round(korea_adjustment * 100, 2),
        "debt_penalty_pct": round(debt_penalty * 100, 2),
        "sector_penalty_pct": round(sector_penalty * 100, 2),
        "beta_used": round(beta, 3),
        "beta_source": beta_source,
    }


# ====================================================================
# 단일 종목 분석
# ====================================================================
def analyze_stock(stock_info: dict) -> dict | None:
    ticker = stock_info["ticker"]
    name = stock_info["name"]
    sector = stock_info.get("sector", "")

    data = fetch_data(ticker)
    if not data or not data.get("current_price"):
        return None

    irrs = calculate_irrs(data)
    rr = calculate_required_return(data, sector)

    primary_irr = irrs.get("primary_irr_pct")
    required = rr.get("required_return_pct")
    margin_of_safety = round(primary_irr - required, 2) if primary_irr is not None and required is not None else None

    # 종합 등급
    if margin_of_safety is None:
        rating = "데이터 부족"
        rating_color = "gray"
    elif margin_of_safety >= 5:
        rating = "강력 매수 후보"
        rating_color = "green"
    elif margin_of_safety >= 2:
        rating = "매수 검토"
        rating_color = "cyan"
    elif margin_of_safety >= 0:
        rating = "시장 평균"
        rating_color = "amber"
    else:
        rating = "고평가"
        rating_color = "red"

    # 강력 매수 후보 조건 (텔레그램 알림 트리거)
    is_strong_buy = (
        primary_irr is not None and primary_irr >= STRONG_BUY_IRR * 100 and
        margin_of_safety is not None and margin_of_safety >= STRONG_BUY_MOS * 100 and
        data.get("roe") is not None and data["roe"] >= STRONG_BUY_ROE
    )

    return {
        "ticker": ticker.replace(".KS", "").replace(".KQ", ""),
        "name": name,
        "sector": sector,
        "market": "KOSPI" if ticker.endswith(".KS") else "KOSDAQ",
        "current_price": data.get("current_price"),
        "market_cap_billion": round(data["market_cap"] / 1e8, 0) if data.get("market_cap") else None,
        "irr": irrs,
        "required_return": rr,
        "margin_of_safety_pct": margin_of_safety,
        "rating": rating,
        "rating_color": rating_color,
        "is_strong_buy": is_strong_buy,
        "snapshot": {
            "per": round(data["trailing_pe"], 2) if data.get("trailing_pe") else None,
            "pbr": round(data["price_to_book"], 2) if data.get("price_to_book") else None,
            "roe_pct": round(data["roe"] * 100, 2) if data.get("roe") else None,
            "dividend_yield_pct": round(data["dividend_yield_raw"], 2) if data.get("dividend_yield_raw") else None,
            "debt_to_equity_pct": round(data["debt_to_equity"], 1) if data.get("debt_to_equity") else None,
        },
    }


# ====================================================================
# 섹터별 집계
# ====================================================================
def aggregate_by_sector(results: list) -> list:
    by_sector = defaultdict(list)
    for r in results:
        if r and r.get("irr", {}).get("primary_irr_pct") is not None:
            by_sector[r["sector"]].append(r["irr"]["primary_irr_pct"])

    summary = []
    for sector, irrs in by_sector.items():
        if not irrs:
            continue
        arr = np.array(irrs)
        summary.append({
            "sector": sector,
            "count": len(irrs),
            "mean_irr_pct": round(float(arr.mean()), 2),
            "median_irr_pct": round(float(np.median(arr)), 2),
            "std_irr_pct": round(float(arr.std()), 2) if len(irrs) > 1 else 0.0,
            "max_irr_pct": round(float(arr.max()), 2),
            "min_irr_pct": round(float(arr.min()), 2),
        })

    summary.sort(key=lambda x: x["mean_irr_pct"], reverse=True)
    return summary


# ====================================================================
# 텔레그램 알림
# ====================================================================
def format_strong_buy_alert(candidates: list) -> str:
    lines = ["📈 IRR 강력 매수 후보 알림", "=" * 30, ""]
    lines.append(f"조건: IRR ≥ {STRONG_BUY_IRR*100:.0f}% + 안전마진 ≥ {STRONG_BUY_MOS*100:.0f}%pt + ROE ≥ {STRONG_BUY_ROE*100:.0f}%")
    lines.append("")

    for c in candidates[:10]:
        snap = c.get("snapshot", {})
        irr = c.get("irr", {}).get("primary_irr_pct")
        mos = c.get("margin_of_safety_pct")
        rr = c.get("required_return", {}).get("required_return_pct")
        lines.append(f"🔸 {c['name']} ({c['ticker']}) — {c['sector']}")
        lines.append(f"   IRR {irr}% / 필요수익률 {rr}% / 안전마진 +{mos}%pt")
        lines.append(f"   PER {snap.get('per', '—')} · PBR {snap.get('pbr', '—')} · ROE {snap.get('roe_pct', '—')}% · 배당 {snap.get('dividend_yield_pct', '—')}%")
        lines.append("")

    if len(candidates) > 10:
        lines.append(f"... 외 {len(candidates) - 10}건 (대시보드 참조)")
        lines.append("")

    lines.append("🚨 통계 추정. 실거래 단독 사용 금지.")
    lines.append("대시보드: https://15678910.github.io/ai-finance/")
    return "\n".join(lines)


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 72)
    print("  한국 주식 IRR 분석기 (다중 방법 + 위험 조정)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  대상: {len(TARGET_STOCKS)}개 종목")
    print("=" * 72)

    state = load_state(STATE_NAME, default={"last_strong_buy_tickers": []})
    last_alerted = set(state.get("last_strong_buy_tickers", []))

    # 종목별 분석
    print("\n[분석 시작]")
    results = []
    for stock in TARGET_STOCKS:
        print(f"  {stock['name']:14s}", end=" ")
        r = analyze_stock(stock)
        if r is None:
            print("→ 데이터 부족")
            continue
        irr = r.get("irr", {}).get("primary_irr_pct")
        mos = r.get("margin_of_safety_pct")
        print(f"→ IRR {irr}% / MoS {mos}%pt / {r['rating']}")
        results.append(r)

    # 정렬: 안전마진 내림차순
    results.sort(key=lambda x: x.get("margin_of_safety_pct") or -999, reverse=True)

    # 섹터별 집계
    sector_summary = aggregate_by_sector(results)

    # 강력 매수 후보
    strong_buys = [r for r in results if r.get("is_strong_buy")]
    new_strong_buys = [r for r in strong_buys if r["ticker"] not in last_alerted]

    # 결과 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "analysis_period": "12개월 trailing 데이터 기준",
        "methodology": "Gordon IRR + Earnings Yield + ROIC-Retention 종합 (중앙값)",
        "assumptions": {
            "risk_free_kr_pct": round(RISK_FREE_KR * 100, 2),
            "equity_premium_pct": round(EQUITY_PREMIUM * 100, 2),
            "korea_discount_pct": round(KOREA_DISCOUNT * 100, 2),
            "growth_cap_high_pct": round(GROWTH_CAP_HIGH * 100, 2),
            "growth_cap_low_pct": round(GROWTH_CAP_LOW * 100, 2),
        },
        "trigger_thresholds": {
            "strong_buy_irr_pct": round(STRONG_BUY_IRR * 100, 2),
            "strong_buy_mos_pct": round(STRONG_BUY_MOS * 100, 2),
            "strong_buy_roe_pct": round(STRONG_BUY_ROE * 100, 2),
        },
        "results": results,
        "sector_summary": sector_summary,
        "strong_buy_candidates_count": len(strong_buys),
        "new_strong_buys_this_run": len(new_strong_buys),
        "warning": "🚨 다중 IRR 방법으로 추정. 통계 모형이며 미래 보장 안 됨. 실거래 단독 사용 금지.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    # 텔레그램 알림 (신규 강력 매수 후보 있을 때만)
    if new_strong_buys:
        msg = format_strong_buy_alert(strong_buys)  # 전체 강력매수 표시, 신규 여부 별도
        try:
            send_message(msg)
            print(f"  ✅ 텔레그램 발송: 강력매수 {len(strong_buys)}건 (신규 {len(new_strong_buys)}건)")
        except Exception as e:
            print(f"  ❌ 텔레그램 발송 실패: {e}")

        # 알림 발송한 ticker 저장 (중복 방지)
        state["last_strong_buy_tickers"] = [r["ticker"] for r in strong_buys]
        save_state(STATE_NAME, state)
    elif strong_buys:
        print(f"  강력매수 후보 {len(strong_buys)}건 (모두 이전 알림 — 미발송)")
    else:
        print("  강력매수 후보 없음")

    # 콘솔 요약
    print("\n[Top 10 안전마진]")
    for r in results[:10]:
        irr = r.get("irr", {}).get("primary_irr_pct")
        mos = r.get("margin_of_safety_pct")
        print(f"  {r['name']:14s} | IRR {irr}% | MoS {mos}%pt | {r['rating']}")

    print("\n[섹터별 IRR (상위 5)]")
    for s in sector_summary[:5]:
        print(f"  {s['sector']:14s} | 평균 {s['mean_irr_pct']}% | 중앙값 {s['median_irr_pct']}% | n={s['count']}")

    print("\n" + "=" * 72)
    print(f"  분석 완료: {len(results)}/{len(TARGET_STOCKS)}개 · 강력매수 {len(strong_buys)}건 · 신규 {len(new_strong_buys)}건")
    print("=" * 72)


if __name__ == "__main__":
    main()
