"""
어닝 리뷰어 (DART API 기반 분기 실적 자동 분석)
================================================

Anthropic Claude for Financial Services의 'Earnings Reviewer Agent'에서 영감을 받아
한국 기업의 분기 실적 발표를 자동으로 추적하고 분석합니다.

수집/분석:
  1. DART API로 분기/연간 보고서 공시 모니터링
  2. 매출/영업이익/순이익 추출
  3. YoY/QoQ 비교
  4. 컨센서스 대비 어닝 서프라이즈 (yfinance EPS 추정치)
  5. 텔레그램 알림 (신규 실적 발표 시)

출력: docs/earnings_reviews.json
🚨 시뮬레이션/분석용. 자동 매매 절대 금지.
"""

import os
import sys
import json
import urllib.request
import urllib.parse
import zipfile
from io import BytesIO
import xml.etree.ElementTree as ET
from datetime import datetime, timezone, timedelta

try:
    import yfinance as yf
except ImportError:
    print("[오류] yfinance 미설치")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "earnings_reviews.json")
STATE_FILE = os.path.join(BASE_DIR, "docs", "earnings_reviews_state.json")
KST = timezone(timedelta(hours=9))

# 추적 종목 (KOSPI 대형주)
TRACKED_STOCKS = [
    {"name": "삼성전자", "stock_code": "005930", "yf": "005930.KS"},
    {"name": "SK하이닉스", "stock_code": "000660", "yf": "000660.KS"},
    {"name": "한미반도체", "stock_code": "042700", "yf": "042700.KS"},
    {"name": "네이버", "stock_code": "035420", "yf": "035420.KS"},
    {"name": "카카오", "stock_code": "035720", "yf": "035720.KS"},
    {"name": "LG에너지솔루션", "stock_code": "373220", "yf": "373220.KS"},
    {"name": "삼성SDI", "stock_code": "006400", "yf": "006400.KS"},
    {"name": "현대차", "stock_code": "005380", "yf": "005380.KS"},
    {"name": "기아", "stock_code": "000270", "yf": "000270.KS"},
    {"name": "POSCO홀딩스", "stock_code": "005490", "yf": "005490.KS"},
    {"name": "셀트리온", "stock_code": "068270", "yf": "068270.KS"},
    {"name": "삼성바이오로직스", "stock_code": "207940", "yf": "207940.KS"},
    {"name": "한화에어로스페이스", "stock_code": "012450", "yf": "012450.KS"},
    {"name": "LIG넥스원", "stock_code": "079550", "yf": "079550.KS"},
    {"name": "현대로템", "stock_code": "064350", "yf": "064350.KS"},
]


# ====================================================================
# DART API
# ====================================================================
class DartAPI:
    BASE = "https://opendart.fss.or.kr/api"

    def __init__(self, api_key: str):
        self.api_key = api_key
        self._corp_codes = None

    def get_corp_codes(self) -> dict:
        """전체 회사 corp_code 매핑 (stock_code → corp_code)."""
        if self._corp_codes:
            return self._corp_codes

        url = f"{self.BASE}/corpCode.xml?crtfc_key={self.api_key}"
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
                corp_code = (elem.findtext("corp_code") or "").strip()
                stock_code = (elem.findtext("stock_code") or "").strip()
                if corp_code and stock_code:
                    mapping[stock_code] = corp_code

            self._corp_codes = mapping
            print(f"  [DART] corp_code 매핑: {len(mapping)}개")
            return mapping
        except Exception as e:
            print(f"  [DART 오류] corp_code: {e}")
            return {}

    def get_recent_disclosures(self, corp_code: str, days_back: int = 30) -> list:
        """최근 N일 공시 목록 조회."""
        bgn_de = (datetime.now() - timedelta(days=days_back)).strftime("%Y%m%d")
        end_de = datetime.now().strftime("%Y%m%d")
        url = (f"{self.BASE}/list.json?crtfc_key={self.api_key}"
               f"&corp_code={corp_code}&bgn_de={bgn_de}&end_de={end_de}"
               f"&page_count=20")
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            if data.get("status") == "000":
                return data.get("list", [])
        except Exception as e:
            print(f"    [공시 조회 실패] {corp_code}: {e}")
        return []

    def get_quarterly_accounts(self, corp_code: str, year: int, period: str = "11013") -> dict:
        """분기 주요계정 조회.
        period: 11011(연간), 11012(반기), 11013(1Q), 11014(3Q)
        """
        url = (f"{self.BASE}/fnlttSinglAcnt.json?crtfc_key={self.api_key}"
               f"&corp_code={corp_code}&bsns_year={year}&reprt_code={period}")
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            if data.get("status") == "000":
                return data
        except Exception as e:
            print(f"    [재무계정 실패] {corp_code} {year}/{period}: {e}")
        return {}


# ====================================================================
# 분석
# ====================================================================
def parse_account_value(text: str) -> float:
    """공시 금액 문자열 → float (단위: 원)."""
    if not text:
        return 0.0
    try:
        return float(str(text).replace(",", "").replace("(", "-").replace(")", ""))
    except (ValueError, TypeError):
        return 0.0


def extract_key_metrics(accounts_data: dict) -> dict:
    """주요계정에서 매출/영업이익/순이익 추출 (당기, 전기, 전전기)."""
    if not accounts_data or not accounts_data.get("list"):
        return {}

    # 우선순위: CFS(연결) > OFS(개별)
    items = accounts_data.get("list", [])
    cfs_items = [i for i in items if i.get("fs_div") == "CFS"]
    target_items = cfs_items if cfs_items else items

    result = {
        "revenue": {},      # 매출
        "operating_profit": {},  # 영업이익
        "net_income": {},   # 당기순이익
        "fs_div": "연결" if cfs_items else "개별",
    }

    metric_map = {
        "ifrs-full_Revenue": "revenue",
        "ifrs-full_OperatingIncomeLoss": "operating_profit",
        "ifrs-full_ProfitLoss": "net_income",
    }

    for item in target_items:
        account_id = item.get("account_id", "")
        if account_id not in metric_map:
            continue
        key = metric_map[account_id]
        # 당기/전기/전전기 금액
        result[key]["current"] = parse_account_value(item.get("thstrm_amount", "0"))
        result[key]["previous_year"] = parse_account_value(item.get("frmtrm_amount", "0"))
        result[key]["two_years_ago"] = parse_account_value(item.get("bfefrmtrm_amount", "0"))
        # 단위 (보통 KRW)
        result[key]["unit"] = item.get("currency", "KRW")

    return result


def calc_yoy_change(current: float, previous: float) -> float:
    """YoY 변화율 (%)"""
    if not previous or previous == 0:
        return 0.0
    return (current - previous) / abs(previous) * 100


def format_amount_korean(amount: float) -> str:
    """원 단위 → 조/억 표기."""
    if abs(amount) >= 1e12:
        return f"{amount / 1e12:.2f}조"
    elif abs(amount) >= 1e8:
        return f"{amount / 1e8:.0f}억"
    else:
        return f"{amount:,.0f}원"


# ====================================================================
# 컨센서스 비교 (yfinance)
# ====================================================================
def get_consensus_eps(yf_ticker: str) -> dict:
    """yfinance에서 EPS 컨센서스 추정치 조회."""
    try:
        t = yf.Ticker(yf_ticker)
        info = t.info or {}
        return {
            "current_quarter_eps_est": info.get("epsCurrentYear"),
            "next_quarter_eps_est": info.get("epsNextYear"),
            "trailing_eps": info.get("trailingEps"),
            "forward_eps": info.get("forwardEps"),
            "earnings_growth": info.get("earningsGrowth"),
            "revenue_growth": info.get("revenueGrowth"),
        }
    except Exception:
        return {}


# ====================================================================
# 어닝 서프라이즈 점수
# ====================================================================
def assess_earnings(metrics: dict, consensus: dict) -> dict:
    """실적 평가 점수 (0-100) + 시그널."""
    revenue_yoy = calc_yoy_change(
        metrics.get("revenue", {}).get("current", 0),
        metrics.get("revenue", {}).get("previous_year", 0)
    )
    opi_yoy = calc_yoy_change(
        metrics.get("operating_profit", {}).get("current", 0),
        metrics.get("operating_profit", {}).get("previous_year", 0)
    )
    net_yoy = calc_yoy_change(
        metrics.get("net_income", {}).get("current", 0),
        metrics.get("net_income", {}).get("previous_year", 0)
    )

    score = 50  # 중립 기준
    signals = []

    # 매출 성장
    if revenue_yoy > 30:
        score += 20
        signals.append(f"매출 YoY +{revenue_yoy:.1f}% (강한 성장)")
    elif revenue_yoy > 10:
        score += 10
        signals.append(f"매출 YoY +{revenue_yoy:.1f}% (안정 성장)")
    elif revenue_yoy < -10:
        score -= 15
        signals.append(f"매출 YoY {revenue_yoy:.1f}% (감소)")

    # 영업이익
    if opi_yoy > 50:
        score += 25
        signals.append(f"영업이익 YoY +{opi_yoy:.1f}% (폭발적)")
    elif opi_yoy > 20:
        score += 15
        signals.append(f"영업이익 YoY +{opi_yoy:.1f}% (강세)")
    elif opi_yoy > 0:
        score += 5
        signals.append(f"영업이익 YoY +{opi_yoy:.1f}%")
    elif opi_yoy < -20:
        score -= 20
        signals.append(f"영업이익 YoY {opi_yoy:.1f}% (악화)")

    # 컨센서스 대비
    earnings_growth = consensus.get("earnings_growth")
    if earnings_growth is not None:
        cg_pct = earnings_growth * 100 if abs(earnings_growth) < 5 else earnings_growth
        if net_yoy > cg_pct + 10:
            score += 15
            signals.append(f"컨센서스 +{cg_pct:.0f}% 대비 어닝 서프라이즈 (실제 +{net_yoy:.0f}%)")
        elif net_yoy < cg_pct - 10:
            score -= 15
            signals.append(f"컨센서스 +{cg_pct:.0f}% 대비 미스 (실제 {net_yoy:+.0f}%)")

    score = max(0, min(100, score))

    if score >= 75:
        rating = "🟢 강력 매수 신호"
    elif score >= 60:
        rating = "🟢 긍정"
    elif score >= 45:
        rating = "🟡 중립"
    elif score >= 30:
        rating = "🟠 부정"
    else:
        rating = "🔴 강력 부정"

    return {
        "score": round(score, 1),
        "rating": rating,
        "signals": signals,
        "revenue_yoy_pct": round(revenue_yoy, 2),
        "operating_profit_yoy_pct": round(opi_yoy, 2),
        "net_income_yoy_pct": round(net_yoy, 2),
    }


# ====================================================================
# 상태 관리
# ====================================================================
def load_state():
    if not os.path.exists(STATE_FILE):
        return {"reviewed": {}}
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"reviewed": {}}


def save_state(state):
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ====================================================================
# 텔레그램
# ====================================================================
def parse_env(env_path: str) -> dict:
    env_vars = {}
    if not os.path.exists(env_path):
        return env_vars
    with open(env_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            v = v.strip().strip("'\"")
            env_vars[k.strip()] = v
    return env_vars


def send_telegram(reviews: list):
    """신규 실적 발표만 알림."""
    new_reviews = [r for r in reviews if r.get("is_new")]
    if not new_reviews:
        return

    env_vars = parse_env(os.path.join(CONFIG_DIR, ".env"))
    bot_token = env_vars.get("TELEGRAM_FINANCE_BOT_TOKEN") or os.environ.get("TELEGRAM_FINANCE_BOT_TOKEN")
    chat_id = env_vars.get("TELEGRAM_FINANCE_CHAT_ID") or os.environ.get("TELEGRAM_FINANCE_CHAT_ID")
    if not bot_token or not chat_id:
        return

    lines = ["📊 분기 실적 발표 알림", "=" * 25, ""]
    for r in new_reviews[:5]:  # 최대 5건
        lines.append(f"\n{r['rating']} {r['name']} ({r['stock_code']})")
        lines.append(f"  {r.get('period_label', '분기')} ({r.get('fs_div', '연결')})")
        if r.get("revenue_current"):
            lines.append(f"  매출: {format_amount_korean(r['revenue_current'])} (YoY {r['revenue_yoy_pct']:+.1f}%)")
        if r.get("operating_profit_current") is not None:
            lines.append(f"  영업이익: {format_amount_korean(r['operating_profit_current'])} (YoY {r['operating_profit_yoy_pct']:+.1f}%)")
        if r.get("net_income_current") is not None:
            lines.append(f"  순이익: {format_amount_korean(r['net_income_current'])} (YoY {r['net_income_yoy_pct']:+.1f}%)")
        for s in r.get("signals", [])[:2]:
            lines.append(f"  • {s}")

    lines.append("\n🚨 시뮬레이션. 자동 매매 금지.")
    lines.append("\n대시보드: https://15678910.github.io/ai-finance/")

    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        body = urllib.parse.urlencode({"chat_id": chat_id, "text": "\n".join(lines)}).encode()
        req = urllib.request.Request(url, data=body, method="POST")
        with urllib.request.urlopen(req, timeout=10) as resp:
            json.loads(resp.read())
        print(f"  [텔레그램] {len(new_reviews)}건 알림 전송")
    except Exception as e:
        print(f"  [텔레그램] 실패: {e}")


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 65)
    print("  어닝 리뷰어 (DART 분기 실적 자동 분석)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    # DART API 키 확인
    api_key = os.environ.get("DART_API_KEY")
    if not api_key:
        env_vars = parse_env(os.path.join(CONFIG_DIR, ".env"))
        api_key = env_vars.get("DART_API_KEY")
    if not api_key:
        print("[오류] DART_API_KEY 미설정. config/.env 또는 환경변수에 설정 필요.")
        sys.exit(1)

    dart = DartAPI(api_key)
    corp_codes = dart.get_corp_codes()
    if not corp_codes:
        print("[오류] DART corp_code 로드 실패")
        sys.exit(1)

    # 현재 연도 + 직전 분기 조회
    now = datetime.now(KST)
    year = now.year

    # 분기 매핑 (보통 발표 시점 기준)
    # 1Q: 5월, 2Q(반기): 8월, 3Q: 11월, 4Q(연간): 3월
    if now.month >= 11:
        period_code = "11014"  # 3Q
        period_label = f"{year}년 3분기"
    elif now.month >= 8:
        period_code = "11012"  # 반기
        period_label = f"{year}년 반기"
    elif now.month >= 5:
        period_code = "11013"  # 1Q
        period_label = f"{year}년 1분기"
    else:
        # 1~4월: 작년 연간 보고서
        year -= 1
        period_code = "11011"
        period_label = f"{year}년 연간"

    print(f"\n[조회 대상] {period_label} (코드: {period_code})")

    state = load_state()
    reviews = []

    for stock in TRACKED_STOCKS:
        stock_code = stock["stock_code"]
        corp_code = corp_codes.get(stock_code)
        if not corp_code:
            print(f"  [{stock['name']}] corp_code 없음 - 건너뜀")
            continue

        print(f"\n  [{stock['name']}] 조회 중...")

        # 주요계정 조회
        accounts = dart.get_quarterly_accounts(corp_code, year, period_code)
        if not accounts.get("list"):
            print(f"    [경고] 데이터 없음")
            continue

        # 메트릭 추출
        metrics = extract_key_metrics(accounts)
        if not metrics.get("revenue"):
            print(f"    [경고] 매출 데이터 추출 실패")
            continue

        # 컨센서스
        consensus = get_consensus_eps(stock["yf"])

        # 평가
        assessment = assess_earnings(metrics, consensus)

        # 신규 발표 여부 (state file 기반)
        review_key = f"{stock_code}_{year}_{period_code}"
        is_new = review_key not in state.get("reviewed", {})
        if is_new:
            state.setdefault("reviewed", {})[review_key] = {
                "reviewed_at": datetime.now(timezone.utc).isoformat(),
                "score": assessment["score"],
            }

        review = {
            "name": stock["name"],
            "stock_code": stock_code,
            "year": year,
            "period_code": period_code,
            "period_label": period_label,
            "fs_div": metrics.get("fs_div", "N/A"),
            "revenue_current": metrics.get("revenue", {}).get("current"),
            "revenue_previous": metrics.get("revenue", {}).get("previous_year"),
            "operating_profit_current": metrics.get("operating_profit", {}).get("current"),
            "operating_profit_previous": metrics.get("operating_profit", {}).get("previous_year"),
            "net_income_current": metrics.get("net_income", {}).get("current"),
            "net_income_previous": metrics.get("net_income", {}).get("previous_year"),
            "consensus": consensus,
            "is_new": is_new,
            **assessment,
        }
        reviews.append(review)

        # 출력
        rev = metrics.get("revenue", {}).get("current", 0)
        opi = metrics.get("operating_profit", {}).get("current", 0)
        net = metrics.get("net_income", {}).get("current", 0)
        print(f"    매출 {format_amount_korean(rev)} (YoY {assessment['revenue_yoy_pct']:+.1f}%)")
        print(f"    영업이익 {format_amount_korean(opi)} (YoY {assessment['operating_profit_yoy_pct']:+.1f}%)")
        print(f"    순이익 {format_amount_korean(net)} (YoY {assessment['net_income_yoy_pct']:+.1f}%)")
        print(f"    {assessment['rating']} (점수 {assessment['score']})")

    # 저장
    save_state(state)
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "period_label": period_label,
        "reviews": sorted(reviews, key=lambda x: x.get("score", 0), reverse=True),
        "warning": "🚨 시뮬레이션. 자동 매매 금지.",
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    # 텔레그램
    send_telegram(reviews)

    print("\n" + "=" * 65)
    print(f"  분석 완료: {len(reviews)}개 종목")
    print("  ⚠️ 시뮬레이션 전용. 자동 매매 절대 금지.")
    print("=" * 65)


if __name__ == "__main__":
    main()
