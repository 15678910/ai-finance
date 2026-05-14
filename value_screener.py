"""
저평가 종목 스크리너 (DART + yfinance 하이브리드)
================================================

DART API와 yfinance를 결합하여 저평가된 한국 주식을 찾아냅니다.

핵심 지표:
  가치 (Valuation): PER, PBR, PSR
  수익성 (Quality): ROE, 영업이익률, 부채비율
  성장 (Growth): 매출 성장률, 영업이익 성장률
  주주환원: 배당수익률

종합 점수 = (가치 점수 × 35%) + (수익성 점수 × 35%) + (성장 점수 × 20%) + (주주환원 × 10%)

🚨 절대 규칙:
  - 시뮬레이션 / 분석용 전용
  - 자동 매매 절대 금지
  - 사용자가 직접 검토 후 투자 결정
"""

import os
import sys
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta
from pathlib import Path

try:
    import yfinance as yf
    import pandas as pd
    import numpy as np
except ImportError as e:
    print(f"[오류] 라이브러리 미설치: {e}")
    print("설치: pip install yfinance pandas numpy")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "value_screener.json")

# ====================================================================
# 분석 대상 종목 (KOSPI 대형주 + KOSDAQ 우량주)
# ====================================================================
TARGET_STOCKS = [
    # IT/반도체 (HBM, AI 메모리, 장비 포함)
    {"ticker": "005930.KS", "name": "삼성전자", "sector": "IT/반도체"},
    {"ticker": "000660.KS", "name": "SK하이닉스", "sector": "IT/반도체"},
    {"ticker": "042700.KS", "name": "한미반도체", "sector": "반도체장비"},
    {"ticker": "000990.KS", "name": "DB하이텍", "sector": "반도체"},
    {"ticker": "035420.KS", "name": "네이버", "sector": "IT"},
    {"ticker": "035720.KS", "name": "카카오", "sector": "IT"},
    {"ticker": "066570.KS", "name": "LG전자", "sector": "전자"},
    {"ticker": "018260.KS", "name": "삼성에스디에스", "sector": "IT"},

    # 금융
    {"ticker": "055550.KS", "name": "신한지주", "sector": "금융"},
    {"ticker": "086790.KS", "name": "하나금융지주", "sector": "금융"},
    {"ticker": "105560.KS", "name": "KB금융", "sector": "금융"},
    {"ticker": "316140.KS", "name": "우리금융지주", "sector": "금융"},

    # 자동차
    {"ticker": "005380.KS", "name": "현대차", "sector": "자동차"},
    {"ticker": "000270.KS", "name": "기아", "sector": "자동차"},
    {"ticker": "012330.KS", "name": "현대모비스", "sector": "자동차"},

    # 화학/에너지
    {"ticker": "051910.KS", "name": "LG화학", "sector": "화학"},
    {"ticker": "010950.KS", "name": "S-Oil", "sector": "에너지"},
    {"ticker": "096770.KS", "name": "SK이노베이션", "sector": "에너지"},
    {"ticker": "015760.KS", "name": "한국전력", "sector": "유틸리티"},

    # 철강/조선
    {"ticker": "005490.KS", "name": "POSCO홀딩스", "sector": "철강"},
    {"ticker": "009540.KS", "name": "HD한국조선해양", "sector": "조선"},

    # 바이오/제약
    {"ticker": "068270.KS", "name": "셀트리온", "sector": "바이오"},
    {"ticker": "207940.KS", "name": "삼성바이오로직스", "sector": "바이오"},
    {"ticker": "128940.KS", "name": "한미약품", "sector": "제약"},
    {"ticker": "000100.KS", "name": "유한양행", "sector": "제약"},

    # 방산
    {"ticker": "012450.KS", "name": "한화에어로스페이스", "sector": "방산"},
    {"ticker": "079550.KS", "name": "LIG넥스원", "sector": "방산"},
    {"ticker": "064350.KS", "name": "현대로템", "sector": "방산"},

    # 배터리/2차전지
    {"ticker": "373220.KS", "name": "LG에너지솔루션", "sector": "배터리"},
    {"ticker": "006400.KS", "name": "삼성SDI", "sector": "배터리"},
    {"ticker": "247540.KQ", "name": "에코프로비엠", "sector": "배터리"},

    # 통신
    {"ticker": "017670.KS", "name": "SK텔레콤", "sector": "통신"},
    {"ticker": "030200.KS", "name": "KT", "sector": "통신"},

    # 유통/소비재
    {"ticker": "271560.KS", "name": "오리온", "sector": "소비재"},
    {"ticker": "097950.KS", "name": "CJ제일제당", "sector": "소비재"},
    {"ticker": "139480.KS", "name": "이마트", "sector": "유통"},
    {"ticker": "023530.KS", "name": "롯데쇼핑", "sector": "유통"},
    {"ticker": "008770.KS", "name": "호텔신라", "sector": "관광"},

    # 건설/물류
    {"ticker": "000720.KS", "name": "현대건설", "sector": "건설"},
    {"ticker": "047040.KS", "name": "대우건설", "sector": "건설"},
    {"ticker": "000120.KS", "name": "CJ대한통운", "sector": "물류"},

    # 게임/엔터
    {"ticker": "036570.KS", "name": "엔씨소프트", "sector": "게임"},
    {"ticker": "259960.KS", "name": "크래프톤", "sector": "게임"},
    {"ticker": "352820.KS", "name": "하이브", "sector": "엔터"},
    {"ticker": "035900.KQ", "name": "JYP엔터테인먼트", "sector": "엔터"},

    # 항공
    {"ticker": "003490.KS", "name": "대한항공", "sector": "항공"},

    # AI/로봇
    {"ticker": "277810.KQ", "name": "레인보우로보틱스", "sector": "로봇"},
]


# ====================================================================
# DART API 호출
# ====================================================================
class DartAPI:
    """OpenDART API 래퍼."""

    BASE_URL = "https://opendart.fss.or.kr/api"

    def __init__(self, api_key: str):
        self.api_key = api_key
        self._corp_code_cache = None

    def get_corp_codes(self) -> dict:
        """전체 상장 회사 corp_code 매핑 (티커 → corp_code).
        KRX의 종목 코드가 아닌 DART 고유 8자리 코드 필요."""
        if self._corp_code_cache:
            return self._corp_code_cache

        import zipfile
        from io import BytesIO
        try:
            from defusedxml import ElementTree as ET  # type: ignore
        except ImportError:
            import xml.etree.ElementTree as ET  # XXE 위험. defusedxml 권장.

        url = f"{self.BASE_URL}/corpCode.xml?crtfc_key={self.api_key}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                content = resp.read()

            with zipfile.ZipFile(BytesIO(content)) as zf:
                with zf.open(zf.namelist()[0]) as xml_file:
                    tree = ET.parse(xml_file)
                    root = tree.getroot()

            mapping = {}
            for elem in root.iter("list"):
                corp_code = elem.findtext("corp_code", "").strip()
                stock_code = elem.findtext("stock_code", "").strip()
                if corp_code and stock_code:
                    mapping[stock_code] = corp_code

            self._corp_code_cache = mapping
            print(f"  [DART] corp_code 매핑 로드: {len(mapping)}개")
            return mapping
        except Exception as e:
            print(f"  [DART 오류] corp_code 로드 실패: {e}")
            return {}

    def get_financial_indicators(self, corp_code: str, year: int) -> dict:
        """특정 회사의 재무지표 조회."""
        result = {}
        # idx_cl_code: M210000 (수익성), M220000 (자본효율성), M230000 (안정성), M240000 (성장성), M250000 (생산성)
        for code, label in [("M210000", "profitability"), ("M220000", "efficiency"),
                             ("M230000", "stability"), ("M240000", "growth")]:
            url = (f"{self.BASE_URL}/fnlttSinglIndx.json?"
                   f"crtfc_key={self.api_key}&corp_code={corp_code}"
                   f"&bsns_year={year}&reprt_code=11011&idx_cl_code={code}")
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=10) as resp:
                    data = json.loads(resp.read().decode("utf-8"))

                if data.get("status") == "000":
                    for item in data.get("list", []):
                        idx_nm = item.get("idx_nm", "")
                        idx_val = item.get("idx_val", "")
                        result[idx_nm] = idx_val
            except Exception:
                pass

        return result


# ====================================================================
# yfinance 데이터 수집
# ====================================================================
def fetch_yfinance_data(ticker: str) -> dict:
    """yfinance에서 가치 지표 수집. PBR 등 누락 시 BVPS로 계산."""
    try:
        stock = yf.Ticker(ticker)
        info = stock.info or {}

        # PBR 폴백 1: priceToBook
        # PBR 폴백 2: 시가총액 / 자기자본총계 (balance sheet)
        price_to_book = info.get("priceToBook")
        current_price = info.get("currentPrice") or info.get("regularMarketPrice")
        market_cap = info.get("marketCap")

        if not price_to_book:
            book_value = info.get("bookValue")
            if current_price and book_value and book_value > 0:
                price_to_book = round(current_price / book_value, 2)

        if not price_to_book and market_cap:
            # balance sheet에서 자기자본총계 추출
            try:
                bs = stock.balance_sheet
                if bs is not None and not bs.empty:
                    # 최신 분기 자기자본 ('Stockholders Equity' 또는 'Total Equity Gross Minority Interest')
                    equity_keys = ['Stockholders Equity', 'Total Equity Gross Minority Interest',
                                   'Total Stockholder Equity', 'Common Stock Equity']
                    equity = None
                    for key in equity_keys:
                        if key in bs.index:
                            equity = bs.loc[key].iloc[0]
                            break
                    if equity and equity > 0:
                        price_to_book = round(market_cap / equity, 2)
            except Exception:
                pass

        return {
            "current_price": info.get("currentPrice") or info.get("regularMarketPrice"),
            "market_cap": info.get("marketCap"),
            "trailing_pe": info.get("trailingPE"),
            "forward_pe": info.get("forwardPE"),
            "price_to_book": price_to_book,
            "price_to_sales": info.get("priceToSalesTrailing12Months"),
            "book_value": info.get("bookValue"),
            "roe": info.get("returnOnEquity"),  # 소수
            "operating_margin": info.get("operatingMargins"),
            "profit_margin": info.get("profitMargins"),
            "debt_to_equity": info.get("debtToEquity"),
            "dividend_yield": info.get("dividendYield"),
            "revenue_growth": info.get("revenueGrowth"),
            "earnings_growth": info.get("earningsGrowth"),
            "beta": info.get("beta"),
            "52w_high": info.get("fiftyTwoWeekHigh"),
            "52w_low": info.get("fiftyTwoWeekLow"),
        }
    except Exception as e:
        print(f"    [yfinance 오류] {ticker}: {e}")
        return {}


# ====================================================================
# 가치 점수 계산
# ====================================================================
# ====================================================================
# yfinance 데이터 형식 헬퍼
# ====================================================================
def normalize_dividend_yield(raw):
    """yfinance dividendYield → 소수 변환 (0~1).

    배경: yfinance 최근 버전(0.2.40+)은 dividendYield를 % 포맷으로 반환
      예: 0.52 = 0.52%, 5.0 = 5.0%
    이전 버전은 decimal 포맷 (0.005 = 0.5%, 0.05 = 5%)
    구분 규칙:
      - raw > 0.3: 비현실적 decimal(30%+ 배당) → % 포맷으로 간주, /100
      - raw <= 0.3: 둘 다 가능. 안전하게 % 포맷 가정 (/100)
        (decimal 가정 시 0.05를 5%로 해석 — 너무 큰 값으로 평가 위험)
    """
    if raw is None:
        return None
    return raw / 100.0


def cap_growth_rate(raw, lo=-0.30, hi=0.50):
    """성장률 캡: 비정상 값 필터.

    이유: yfinance의 earningsGrowth는 전년 동기 대비 — 손익 적자/흑자 전환 시
      비정상적으로 큰 값(예: 4.96 = +496%) 또는 음수 폭발 발생.
    """
    if raw is None:
        return None
    if raw > hi or raw < lo:
        return None  # 신뢰 불가
    return raw


# ====================================================================
# IRR (Internal Rate of Return) 추정
# ====================================================================
# 한국 국고채 10년 수익률 기준 (변동 시 갱신 가능). 무위험금리 프록시.
RISK_FREE_RATE_KR = 0.035

def calculate_irr_metrics(yf_data: dict) -> dict:
    """주식 IRR (기대 연수익률) 추정.

    3가지 방법으로 계산 후 가장 적합한 값을 primary로 선택:
      1. Gordon Growth IRR: IRR = 배당수익률 + 영구성장률 (배당주 적합)
      2. Earnings Yield: IRR ≈ 1/PER (무배당/저배당 종목 프록시)
      3. 종합 추정: 배당주는 Gordon, 무배당은 Earnings Yield

    영구성장률(g) 추정:
      - revenue_growth / earnings_growth 평균
      - 8% 상단 캡 (GDP 성장률 한도), -5% 하단 플로어 (회복 가능)
      - 데이터 부족 시 ROE × (1 - 배당성향) 의 지속가능성장률 모형
      - 모두 없으면 디폴트 3%

    한국 시장 특수성:
      - 코리아 디스카운트 ~30~40% → 위험 프리미엄 상향 권장 (+1~2%p)
      - 배당성향 평균 25% (미국 40%+ 대비 낮음)
      - 본 함수는 Raw IRR 반환, 디스카운트 보정은 호출부에서 처리
    """
    # yfinance dividendYield는 최근 버전에서 % 포맷 (0.52 = 0.52%)
    div_yield = normalize_dividend_yield(yf_data.get("dividend_yield"))

    trailing_pe = yf_data.get("trailing_pe")
    forward_pe = yf_data.get("forward_pe")
    pe = trailing_pe or forward_pe

    # 성장률 캡 (저기저 효과로 인한 비정상 값 필터)
    revenue_growth = cap_growth_rate(yf_data.get("revenue_growth"))
    earnings_growth = cap_growth_rate(yf_data.get("earnings_growth"))
    roe = yf_data.get("roe")

    # ─── 영구성장률(g) 추정 ───────────────────────────
    growth_candidates = []
    if revenue_growth is not None:
        growth_candidates.append(revenue_growth)
    if earnings_growth is not None:
        growth_candidates.append(earnings_growth)

    if growth_candidates:
        avg_growth = sum(growth_candidates) / len(growth_candidates)
        # 한국 명목 GDP 성장률 ~4-5% 수준 → 영구 8% 상단 캡
        # (단기 호황 cyclical 효과를 영구로 가정 금지)
        g = max(min(avg_growth, 0.08), -0.05)
        g_source = "revenue/earnings growth (capped 8%)"
    elif roe is not None and div_yield is not None and pe and pe > 0:
        # 지속가능성장률 모델: g = ROE × (1 - 배당성향)
        payout_ratio = min(max(div_yield * pe, 0), 1.0)
        g = max(min(roe * (1 - payout_ratio), 0.08), 0.01)
        g_source = "sustainable growth (ROE × retention)"
    else:
        g = 0.03
        g_source = "default 3%"

    # ─── Gordon Growth IRR ──────────────────────────
    gordon_irr = None
    if div_yield and div_yield > 0.001:  # 0.1% 이상 배당 시
        gordon_irr = div_yield + g

    # ─── Earnings Yield (PER 역수) ───────────────────
    earnings_yield = None
    if pe and pe > 0:
        earnings_yield = 1 / pe

    # ─── 종합 (primary IRR 선택) ─────────────────────
    if gordon_irr is not None:
        primary_irr = gordon_irr
        method = "gordon"
    elif earnings_yield is not None:
        primary_irr = earnings_yield
        method = "earnings_yield"
    else:
        primary_irr = None
        method = None

    # 무위험금리 대비 스프레드
    spread = (primary_irr - RISK_FREE_RATE_KR) if primary_irr is not None else None

    return {
        "gordon_irr_pct": round(gordon_irr * 100, 2) if gordon_irr is not None else None,
        "earnings_yield_pct": round(earnings_yield * 100, 2) if earnings_yield is not None else None,
        "implied_growth_pct": round(g * 100, 2),
        "growth_source": g_source,
        "primary_irr_pct": round(primary_irr * 100, 2) if primary_irr is not None else None,
        "irr_method": method,
        "risk_premium_pct": round(spread * 100, 2) if spread is not None else None,
        "risk_free_rate_pct": round(RISK_FREE_RATE_KR * 100, 2),
    }


def score_per(per: float) -> float:
    """PER 점수 (낮을수록 높은 점수)."""
    if per is None or per <= 0:
        return 0
    if per < 5:
        return 100
    elif per < 10:
        return 80
    elif per < 15:
        return 60
    elif per < 20:
        return 40
    elif per < 30:
        return 20
    else:
        return 5


def score_pbr(pbr: float) -> float:
    """PBR 점수 (낮을수록 높은 점수)."""
    if pbr is None or pbr <= 0:
        return 0
    if pbr < 0.5:
        return 100
    elif pbr < 1.0:
        return 85
    elif pbr < 1.5:
        return 65
    elif pbr < 2.0:
        return 45
    elif pbr < 3.0:
        return 25
    else:
        return 10


def score_roe(roe: float) -> float:
    """ROE 점수 (높을수록 높은 점수). roe는 소수 (0.15 = 15%)."""
    if roe is None:
        return 0
    roe_pct = roe * 100 if abs(roe) < 5 else roe  # 소수면 *100
    if roe_pct < 0:
        return 0
    elif roe_pct < 5:
        return 20
    elif roe_pct < 10:
        return 50
    elif roe_pct < 15:
        return 75
    elif roe_pct < 20:
        return 90
    else:
        return 100


def score_debt(debt_ratio: float) -> float:
    """부채비율 점수 (낮을수록 높은 점수). debt_ratio는 % 단위."""
    if debt_ratio is None or debt_ratio < 0:
        return 50  # 데이터 없으면 중립
    if debt_ratio < 50:
        return 100
    elif debt_ratio < 100:
        return 70
    elif debt_ratio < 150:
        return 40
    elif debt_ratio < 200:
        return 20
    else:
        return 5


def score_growth(growth: float) -> float:
    """성장률 점수 (높을수록 높은 점수). growth는 소수."""
    if growth is None:
        return 30  # 데이터 없으면 중립 약간
    growth_pct = growth * 100 if abs(growth) < 5 else growth
    if growth_pct < -10:
        return 0
    elif growth_pct < 0:
        return 20
    elif growth_pct < 5:
        return 50
    elif growth_pct < 10:
        return 70
    elif growth_pct < 20:
        return 85
    else:
        return 100


def score_dividend(div_yield_pct: float) -> float:
    """배당수익률 점수 (높을수록 높은 점수).

    입력: 배당수익률 (% 단위, yfinance 0.2.40+ 포맷)
      예: 0.52 = 0.52%, 5.0 = 5.0%
    """
    if div_yield_pct is None or div_yield_pct < 0:
        return 30
    if div_yield_pct < 1:
        return 20
    elif div_yield_pct < 2:
        return 50
    elif div_yield_pct < 4:
        return 80
    elif div_yield_pct < 6:
        return 95
    else:
        return 100


def calculate_value_score(metrics: dict) -> dict:
    """종합 가치 점수 계산.

    종합 점수 = (가치 35%) + (수익성 35%) + (성장 20%) + (주주환원 10%)
    """
    # trailingPE 없으면 forwardPE 사용
    per_val = metrics.get("trailing_pe") or metrics.get("forward_pe")
    s_per = score_per(per_val)
    s_pbr = score_pbr(metrics.get("price_to_book"))
    s_roe = score_roe(metrics.get("roe"))
    s_debt = score_debt(metrics.get("debt_to_equity"))
    s_growth = score_growth(metrics.get("revenue_growth"))
    s_dividend = score_dividend(metrics.get("dividend_yield"))

    # 카테고리 점수
    valuation = (s_per * 0.5) + (s_pbr * 0.5)
    quality = (s_roe * 0.6) + (s_debt * 0.4)
    growth = s_growth
    shareholder = s_dividend

    # 종합 점수
    total = (valuation * 0.35) + (quality * 0.35) + (growth * 0.20) + (shareholder * 0.10)

    return {
        "total_score": round(total, 1),
        "valuation_score": round(valuation, 1),
        "quality_score": round(quality, 1),
        "growth_score": round(growth, 1),
        "shareholder_score": round(shareholder, 1),
        "breakdown": {
            "per_score": round(s_per, 1),
            "pbr_score": round(s_pbr, 1),
            "roe_score": round(s_roe, 1),
            "debt_score": round(s_debt, 1),
            "growth_score": round(s_growth, 1),
            "dividend_score": round(s_dividend, 1),
        },
    }


def value_trap_filter(metrics: dict) -> tuple:
    """저평가 함정 필터링. (통과 여부, 사유)"""
    # 시가총액 1,000억원 이하 제외
    market_cap = metrics.get("market_cap")
    if market_cap and market_cap < 100_000_000_000:
        return False, f"시가총액 {market_cap/1e8:.0f}억 미만 (1000억 기준)"

    # ROE 매우 음수 (-20% 미만) 제외 - 일시적 적자는 허용
    roe = metrics.get("roe")
    if roe is not None and roe < -0.20:
        return False, f"ROE {roe*100:.1f}% (심각한 적자)"

    # 부채비율 500% 초과 제외 (완화: 200→500)
    debt = metrics.get("debt_to_equity")
    if debt and debt > 500:
        return False, f"부채비율 {debt:.0f}% 과다"

    return True, "통과"


# ====================================================================
# 종목 분석
# ====================================================================
def analyze_stock(stock_info: dict, dart_api: DartAPI = None) -> dict:
    """단일 종목 분석."""
    ticker = stock_info["ticker"]
    name = stock_info["name"]

    print(f"  [{name}] 분석 중...")

    # yfinance 데이터
    yf_data = fetch_yfinance_data(ticker)
    if not yf_data or not yf_data.get("current_price"):
        return None

    # 점수 계산
    score = calculate_value_score(yf_data)

    # 함정 필터
    passed, reason = value_trap_filter(yf_data)

    # IRR (기대 연수익률) 추정
    irr = calculate_irr_metrics(yf_data)

    # 결과 조립
    result = {
        "ticker": ticker.replace(".KS", "").replace(".KQ", ""),
        "name": name,
        "sector": stock_info.get("sector", ""),
        "market": "KOSPI" if ticker.endswith(".KS") else "KOSDAQ",
        "current_price": yf_data.get("current_price"),
        "market_cap_billion": round(yf_data["market_cap"] / 1e8, 0) if yf_data.get("market_cap") else None,
        "metrics": {
            "per": round(yf_data["trailing_pe"], 2) if yf_data.get("trailing_pe") else (round(yf_data["forward_pe"], 2) if yf_data.get("forward_pe") else None),
            "forward_per": round(yf_data["forward_pe"], 2) if yf_data.get("forward_pe") else None,
            "pbr": round(yf_data["price_to_book"], 2) if yf_data.get("price_to_book") else None,
            "psr": round(yf_data["price_to_sales"], 2) if yf_data.get("price_to_sales") else None,
            "roe_pct": round(yf_data["roe"] * 100, 2) if yf_data.get("roe") else None,
            "operating_margin_pct": round(yf_data["operating_margin"] * 100, 2) if yf_data.get("operating_margin") else None,
            "debt_to_equity": round(yf_data["debt_to_equity"], 0) if yf_data.get("debt_to_equity") else None,
            "dividend_yield_pct": round(yf_data["dividend_yield"], 2) if yf_data.get("dividend_yield") is not None else None,
            "revenue_growth_pct": round(yf_data["revenue_growth"] * 100, 2) if yf_data.get("revenue_growth") else None,
        },
        "irr": irr,
        "scores": score,
        "filter_passed": passed,
        "filter_reason": reason,
    }

    return result


# ====================================================================
# 텔레그램 전송
# ====================================================================
def send_telegram(top_picks: list):
    """저평가 Top 10 결과를 텔레그램으로 알림."""
    env_path = os.path.join(CONFIG_DIR, ".env")
    bot_token = None
    chat_id = None

    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" not in line or line.startswith("#"):
                    continue
                k, v = line.split("=", 1)
                k, v = k.strip(), v.strip().strip("'\"")
                if k == "TELEGRAM_FINANCE_BOT_TOKEN":
                    bot_token = v
                elif k == "TELEGRAM_FINANCE_CHAT_ID":
                    chat_id = v

    bot_token = bot_token or os.environ.get("TELEGRAM_FINANCE_BOT_TOKEN")
    chat_id = chat_id or os.environ.get("TELEGRAM_FINANCE_CHAT_ID")

    if not bot_token or not chat_id:
        print("  [텔레그램] 설정 없음. 전송 생략.")
        return

    lines = ["🔍 저평가 종목 Top 10", "=" * 25, ""]
    for i, stock in enumerate(top_picks[:10], 1):
        m = stock.get("metrics", {})
        per = m.get("per", "N/A")
        pbr = m.get("pbr", "N/A")
        roe = m.get("roe_pct", "N/A")
        score = stock["scores"]["total_score"]
        lines.append(f"{i}. {stock['name']} ({stock['ticker']}) - {score}점")
        lines.append(f"   PER {per} / PBR {pbr} / ROE {roe}%")
        lines.append("")

    lines.append("🚨 시뮬레이션/분석용. 자동 매매 금지.")
    lines.append(f"\n대시보드: https://15678910.github.io/ai-finance/")

    message = "\n".join(lines)
    if len(message) > 4000:
        message = message[:4000] + "\n... (대시보드 참조)"

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
# 메인
# ====================================================================
def main():
    import argparse
    parser = argparse.ArgumentParser(description="저평가 종목 스크리너")
    parser.add_argument("--no-telegram", action="store_true", help="텔레그램 전송 생략")
    args = parser.parse_args()

    print("=" * 65)
    print("  저평가 종목 스크리너")
    print(f"  대상: {len(TARGET_STOCKS)}개 종목")
    print(f"  시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    # DART API 초기화 (선택)
    dart_api_key = os.environ.get("DART_API_KEY")
    dart_api = None
    if dart_api_key:
        print("  [DART] API 키 확인. 보조 데이터 수집 활성화.")
        dart_api = DartAPI(dart_api_key)
    else:
        print("  [DART] API 키 미설정. yfinance만 사용.")

    # 종목별 분석
    print("\n[분석 시작]")
    results = []
    for stock in TARGET_STOCKS:
        try:
            r = analyze_stock(stock, dart_api)
            if r:
                results.append(r)
            time.sleep(0.3)  # rate limit
        except Exception as e:
            print(f"    [오류] {stock['name']}: {e}")

    # 함정 필터 통과 + 점수 정렬
    passed = [r for r in results if r["filter_passed"]]
    rejected = [r for r in results if not r["filter_passed"]]
    passed.sort(key=lambda x: x["scores"]["total_score"], reverse=True)

    top_10 = passed[:10]

    print(f"\n[결과] 분석 완료: {len(results)}개")
    print(f"        필터 통과: {len(passed)}개")
    print(f"        필터 거부: {len(rejected)}개")
    print(f"        Top 10 추출")

    # 출력
    print(f"\n[저평가 Top 10]")
    for i, stock in enumerate(top_10, 1):
        m = stock["metrics"]
        print(f"  {i:2d}. {stock['name']:15s} ({stock['ticker']}) "
              f"- 점수 {stock['scores']['total_score']:5.1f} "
              f"- PER {m.get('per', 'N/A')} PBR {m.get('pbr', 'N/A')} "
              f"ROE {m.get('roe_pct', 'N/A')}%")

    # 사용자 관심 종목 (필터/순위 무관 항상 표시)
    WATCHLIST = ["042700", "006400"]  # 한미반도체, 삼성SDI
    watchlist_results = [r for r in results if r["ticker"] in WATCHLIST]

    # 저장
    output = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_analyzed": len(results),
        "filter_passed": len(passed),
        "filter_rejected": len(rejected),
        "top_picks": top_10,
        "watchlist": watchlist_results,  # 관심 종목
        "all_passed": passed,  # 전체 통과 종목 (대시보드용)
        "scoring_formula": {
            "valuation_weight": 0.35,
            "quality_weight": 0.35,
            "growth_weight": 0.20,
            "shareholder_weight": 0.10,
        },
        "warning": "🚨 시뮬레이션/분석용. 자동 매매 절대 금지. 사용자 직접 검토 후 투자 결정.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    # 텔레그램
    if not args.no_telegram:
        send_telegram(top_10)

    print("\n" + "=" * 65)
    print("  ⚠️ 본 결과는 시뮬레이션 전용입니다.")
    print("  ⚠️ 실제 투자 결정은 본인의 판단 필요.")
    print("=" * 65)

    return 0


if __name__ == "__main__":
    sys.exit(main())
