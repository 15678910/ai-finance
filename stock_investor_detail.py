"""
종목별 일별 투자자 매매 동향 + 이벤트 분석
================================================

목적:
  사용자가 특정 종목(예: 이마트, 신세계)의 일별 외국인/기관/개인 매매를
  60거래일까지 추적하고, 특정 이벤트(예: 2026-05-23 스타벅스 PR 사건)가
  주가·매매 흐름에 미친 영향을 정량적으로 분석.

데이터:
  · Naver 금융 frgn.naver 페이지 3개 스크래핑 (총 60거래일)
  · 컬럼: 날짜·종가·등락률·거래량·기관 순매매·외국인 순매매·외인 보유율
  · 개인 매매: 추정값 (총 거래량의 부호 반전된 외인+기관 합산)

자동 이벤트 감지:
  · 거래량 200%+ 폭증 (60일 평균 대비)
  · 등락률 -3% 이하 급락 (또는 +3% 이상 급등)
  · 외인 매매 +/-300K주 이상 이상치

이벤트 분석:
  · T-5 ~ T-1 (전 5일 평균)
  · T (이벤트 당일)
  · T+1 ~ T+5 (후 5일 평균)
  · T+6 ~ T+20 (후 4주 평균)
  · 추세 변화 감지 (매수 → 매도 전환 등)

🚨 시뮬레이션. 자동 매매 금지.
"""

import os
import sys
import re
import json
import time
import urllib.request
from datetime import datetime, timezone, timedelta
from statistics import mean, stdev

from core import send_message, load_state, save_state

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "stock_investor_details.json")
STATE_NAME = "stock_investor_details"
KST = timezone(timedelta(hours=9))

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
TIMEOUT_SEC = 15
PAGES_TO_FETCH = 3   # 페이지당 20일 × 3 = 60거래일

# ====================================================================
# 추적 대상 종목 (사용자가 자주 조회할 종목 + 동종업종 비교용)
# ====================================================================
TARGET_STOCKS = [
    # 유통 - 스타벅스 모회사 + 그룹 + 동종업종
    {"ticker": "139480", "name": "이마트",      "category": "유통",
     "context": "스타벅스코리아 67.5% 모회사 (2021년 인수)"},
    {"ticker": "004170", "name": "신세계",      "category": "유통",
     "context": "신세계그룹 (이마트 모기업, 정유경)"},
    {"ticker": "023530", "name": "롯데쇼핑",    "category": "유통",
     "context": "동종업종 비교"},
    {"ticker": "008770", "name": "호텔신라",    "category": "유통",
     "context": "동종업종 (관광·면세)"},
    {"ticker": "069960", "name": "현대백화점",  "category": "유통",
     "context": "동종업종"},
    {"ticker": "057050", "name": "현대홈쇼핑",  "category": "유통",
     "context": "동종업종 (홈쇼핑)"},
    # 사용자가 자주 보는 핵심 종목들 추가
    {"ticker": "005930", "name": "삼성전자",    "category": "IT/반도체",
     "context": "한국 시가총액 1위"},
    {"ticker": "000660", "name": "SK하이닉스",  "category": "IT/반도체",
     "context": "HBM/AI 메모리"},
    {"ticker": "042700", "name": "한미반도체",  "category": "반도체장비",
     "context": "HBM TC 본더"},
    {"ticker": "035420", "name": "네이버",      "category": "IT/플랫폼",
     "context": "AI/검색"},
    {"ticker": "035720", "name": "카카오",      "category": "IT/플랫폼",
     "context": "AI/메신저"},
    {"ticker": "277810", "name": "레인보우로보틱스", "category": "로봇",
     "context": "휴머노이드 로봇 대장주 (삼성전자 인수)"},
]


# ====================================================================
# Naver 페이지 스크래핑 + 60거래일 통합
# ====================================================================
def _parse_int(s: str) -> int:
    if not s or s in ("-", "--", "N/A"):
        return 0
    s = s.replace(",", "").replace("+", "").strip()
    try:
        return int(float(s))
    except Exception:
        return 0


def _parse_pct(s: str) -> float:
    if not s or s in ("-", "--"):
        return 0.0
    s = s.replace("%", "").replace("+", "").replace(",", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0


def fetch_naver_page(ticker: str, page: int = 1) -> list:
    """단일 페이지 (~20일) 스크래핑."""
    url = f"https://finance.naver.com/item/frgn.naver?code={ticker}&page={page}"
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": USER_AGENT,
            "Referer": "https://finance.naver.com/",
        })
        with urllib.request.urlopen(req, timeout=TIMEOUT_SEC) as resp:  # nosec
            html = resp.read().decode("euc-kr", errors="replace")
    except Exception as e:
        print(f"    [실패] {ticker} page {page}: {e}")
        return []

    tr_pattern = re.compile(r"<tr[^>]*>(.*?)</tr>", re.DOTALL)
    date_pattern = re.compile(r"(\d{4}\.\d{2}\.\d{2})")
    rows = []

    for tr in tr_pattern.findall(html):
        if not date_pattern.search(tr):
            continue
        td_texts = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.DOTALL)
        cleaned = []
        for td in td_texts:
            text = re.sub(r"<[^>]+>", "", td)
            text = re.sub(r"\s+", " ", text).strip()
            cleaned.append(text)
        if len(cleaned) < 9:
            continue
        try:
            row = {
                "date": cleaned[0].replace(".", "-"),
                "close": _parse_int(cleaned[1]),
                "change_pct": _parse_pct(cleaned[3]),
                "volume": _parse_int(cleaned[4]),
                "institutional_net_shares": _parse_int(cleaned[5]),
                "foreign_net_shares": _parse_int(cleaned[6]),
                "foreign_holding_shares": _parse_int(cleaned[7]),
                "foreign_holding_pct": _parse_pct(cleaned[8]),
            }
            # 개인 매매 추정: 거래량 부호 (단순화)
            # 정확한 계산은 매수·매도 별도 필요하지만, 순매매 부호로 근사
            row["individual_net_shares_est"] = -(row["foreign_net_shares"] + row["institutional_net_shares"])
            rows.append(row)
        except Exception:
            continue
    return rows


def fetch_stock_60d(ticker: str) -> list:
    """3페이지 통합 (60거래일)."""
    all_rows = []
    seen_dates = set()
    for page in range(1, PAGES_TO_FETCH + 1):
        rows = fetch_naver_page(ticker, page)
        for r in rows:
            if r["date"] not in seen_dates:
                all_rows.append(r)
                seen_dates.add(r["date"])
        time.sleep(0.3)  # rate limit
    # 최신순 정렬
    all_rows.sort(key=lambda r: r["date"], reverse=True)
    return all_rows


# ====================================================================
# 자동 이벤트 감지
# ====================================================================
def detect_auto_events(rows: list) -> list:
    """비정상 거래량·등락률 일자 자동 감지."""
    if len(rows) < 10:
        return []

    volumes = [r["volume"] for r in rows if r["volume"] > 0]
    if not volumes:
        return []
    avg_vol = mean(volumes)
    std_vol = stdev(volumes) if len(volumes) > 1 else 0

    foreign_abs = [abs(r["foreign_net_shares"]) for r in rows]
    avg_foreign = mean(foreign_abs) if foreign_abs else 0

    events = []
    for r in rows:
        reasons = []
        if r["volume"] > avg_vol * 3:
            ratio = r["volume"] / avg_vol if avg_vol > 0 else 0
            reasons.append(f"거래량 +{(ratio-1)*100:.0f}% 폭증")
        if r["change_pct"] <= -3.0:
            reasons.append(f"등락 {r['change_pct']:.2f}% 급락")
        elif r["change_pct"] >= 3.0:
            reasons.append(f"등락 +{r['change_pct']:.2f}% 급등")
        if abs(r["foreign_net_shares"]) > avg_foreign * 3 and avg_foreign > 0:
            reasons.append(f"외인 {r['foreign_net_shares']:+,}주 이상치")

        if reasons:
            events.append({
                "date": r["date"],
                "type": "auto_detected",
                "reasons": reasons,
                "close": r["close"],
                "change_pct": r["change_pct"],
                "volume": r["volume"],
                "foreign_net_shares": r["foreign_net_shares"],
                "institutional_net_shares": r["institutional_net_shares"],
            })
    return events


# ====================================================================
# 이벤트 전후 통계 비교
# ====================================================================
def _stats(rows: list) -> dict:
    if not rows:
        return {"n": 0}
    return {
        "n": len(rows),
        "avg_change_pct": round(mean([r["change_pct"] for r in rows]), 3),
        "total_change_pct": round(sum([r["change_pct"] for r in rows]), 3),
        "avg_volume": int(mean([r["volume"] for r in rows])),
        "total_foreign_net": sum([r["foreign_net_shares"] for r in rows]),
        "total_institutional_net": sum([r["institutional_net_shares"] for r in rows]),
        "total_individual_net_est": sum([r["individual_net_shares_est"] for r in rows]),
        "avg_foreign_holding_pct": round(mean([r["foreign_holding_pct"] for r in rows]), 3),
    }


def analyze_event_window(rows: list, event_date: str,
                         before: int = 5, after: int = 5, after_long: int = 20) -> dict:
    """이벤트 전후 통계 비교 + 추세 변화 감지.

    이벤트일이:
      · 데이터 범위보다 미래 → before-only 분석 (사전 동향)
      · 데이터 범위에 존재 → 정상 전후 비교
      · 데이터 범위보다 과거 → 사전 정의된 윈도우로 분석
    """
    if not rows:
        return {"error": "데이터 없음"}

    # 시간순 정렬
    sorted_rows = sorted(rows, key=lambda r: r["date"])
    latest_date = sorted_rows[-1]["date"]
    earliest_date = sorted_rows[0]["date"]

    # 이벤트일이 미래 (데이터 범위보다 큼) → before-only 분석
    if event_date > latest_date:
        before_rows = sorted_rows[-before:]
        return {
            "event_date_target": event_date,
            "event_date_actual": None,
            "status": "future_or_no_data_yet",
            "message": f"이벤트일({event_date}) 거래 데이터 아직 수집 안 됨. "
                       f"최근 {len(before_rows)}거래일 사전 동향 분석.",
            "latest_data_date": latest_date,
            "earliest_data_date": earliest_date,
            "before_5d": _stats(before_rows),
        }

    # 이벤트 인덱스 (정확 일치 또는 가장 가까운 미래 거래일)
    event_idx = None
    for i, r in enumerate(sorted_rows):
        if r["date"] >= event_date:
            event_idx = i
            break
    if event_idx is None:
        return {"error": "이벤트 일자 데이터 없음"}

    event_day = sorted_rows[event_idx]
    before_rows = sorted_rows[max(0, event_idx - before):event_idx]
    after_rows = sorted_rows[event_idx + 1:event_idx + 1 + after]
    after_long_rows = sorted_rows[event_idx + 1:event_idx + 1 + after_long]

    before_stats = _stats(before_rows)
    event_stats = _stats([event_day])
    after_stats = _stats(after_rows)
    after_long_stats = _stats(after_long_rows)

    # 추세 변화 감지
    trend_changes = []
    if before_stats.get("total_foreign_net", 0) > 0 and after_stats.get("total_foreign_net", 0) < 0:
        trend_changes.append("외국인: 매수 → 매도 전환")
    elif before_stats.get("total_foreign_net", 0) < 0 and after_stats.get("total_foreign_net", 0) > 0:
        trend_changes.append("외국인: 매도 → 매수 전환")
    if before_stats.get("total_institutional_net", 0) > 0 and after_stats.get("total_institutional_net", 0) < 0:
        trend_changes.append("기관: 매수 → 매도 전환")
    elif before_stats.get("total_institutional_net", 0) < 0 and after_stats.get("total_institutional_net", 0) > 0:
        trend_changes.append("기관: 매도 → 매수 전환")

    # 거래량 이상치
    avg_vol = before_stats.get("avg_volume", 1)
    if avg_vol > 0 and event_day["volume"] > avg_vol * 2:
        trend_changes.append(f"이벤트 당일 거래량 +{(event_day['volume']/avg_vol - 1)*100:.0f}% 폭증")

    # 보유율 변화
    if before_stats.get("avg_foreign_holding_pct") and after_stats.get("avg_foreign_holding_pct"):
        holding_chg = after_stats["avg_foreign_holding_pct"] - before_stats["avg_foreign_holding_pct"]
        if abs(holding_chg) >= 0.3:
            direction = "감소" if holding_chg < 0 else "증가"
            trend_changes.append(f"외인 보유율 {holding_chg:+.2f}%p {direction}")

    return {
        "event_date_target": event_date,
        "event_date_actual": event_day["date"],
        "event_day": {
            "close": event_day["close"],
            "change_pct": event_day["change_pct"],
            "volume": event_day["volume"],
            "foreign_net_shares": event_day["foreign_net_shares"],
            "institutional_net_shares": event_day["institutional_net_shares"],
            "individual_net_shares_est": event_day["individual_net_shares_est"],
            "foreign_holding_pct": event_day["foreign_holding_pct"],
        },
        "before_5d": before_stats,
        "after_5d": after_stats,
        "after_20d": after_long_stats,
        "trend_changes": trend_changes,
    }


# ====================================================================
# 메인 — 종목별 수집 + 자동 이벤트 + 사전 정의 이벤트 분석
# ====================================================================
# 사전 정의 이벤트 (분석 대상)
PREDEFINED_EVENTS = []   # 사전 정의 이벤트 제거 — 기업 기술 프로필로 대체


def main():
    print("=" * 72)
    print("  종목별 일별 투자자 매매 동향 + 이벤트 분석")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  대상: {len(TARGET_STOCKS)}개 종목 × {PAGES_TO_FETCH * 20}거래일")
    print("=" * 72)

    state = load_state(STATE_NAME, default={"events": [], "user_events": []})

    print("\n[종목별 60일 데이터 수집]")
    stocks_data = []
    for stock in TARGET_STOCKS:
        ticker = stock["ticker"]
        print(f"  {stock['name']:14s} ({ticker})...", end=" ")
        rows = fetch_stock_60d(ticker)
        if not rows:
            print("실패")
            continue

        # 자동 이벤트 감지
        auto_events = detect_auto_events(rows)

        # 사전 정의 이벤트 분석 (해당 종목인 경우)
        event_analyses = []
        for evt in PREDEFINED_EVENTS:
            if ticker in evt.get("affected_tickers", []) or \
               ticker in evt.get("comparison_tickers", []):
                analysis = analyze_event_window(rows, evt["date"])
                if "error" not in analysis:
                    event_analyses.append({
                        "label": evt["label"],
                        "description": evt["description"],
                        "is_affected": ticker in evt.get("affected_tickers", []),
                        "is_comparison": ticker in evt.get("comparison_tickers", []),
                        **analysis,
                    })

        stocks_data.append({
            "ticker": ticker,
            "name": stock["name"],
            "category": stock["category"],
            "context": stock.get("context", ""),
            "days_collected": len(rows),
            "latest_date": rows[0]["date"] if rows else None,
            "earliest_date": rows[-1]["date"] if rows else None,
            "rows": rows,
            "auto_detected_events": auto_events,
            "predefined_event_analyses": event_analyses,
        })
        print(f"{len(rows)}일 + 자동 이벤트 {len(auto_events)}건")

    # 사전 정의 이벤트 종합 리포트
    print("\n[사전 정의 이벤트 종합 분석]")
    event_reports = []
    for evt in PREDEFINED_EVENTS:
        print(f"  📅 {evt['label']} ({evt['date']})")
        affected_results = []
        comparison_results = []
        for sd in stocks_data:
            for ea in sd.get("predefined_event_analyses", []):
                if ea["label"] != evt["label"]:
                    continue
                summary = {
                    "ticker": sd["ticker"],
                    "name": sd["name"],
                    "context": sd["context"],
                    "status": ea.get("status"),
                    "message": ea.get("message"),
                    "before_5d": ea.get("before_5d", {}),
                    "after_5d": ea.get("after_5d", {}),
                    "trend_changes": ea.get("trend_changes", []),
                }
                # 이벤트 당일 데이터가 있는 경우만
                if ea.get("event_day"):
                    summary.update({
                        "event_day_close": ea["event_day"].get("close"),
                        "event_day_change_pct": ea["event_day"].get("change_pct"),
                        "event_day_volume": ea["event_day"].get("volume"),
                        "event_day_foreign": ea["event_day"].get("foreign_net_shares"),
                        "event_day_institutional": ea["event_day"].get("institutional_net_shares"),
                    })

                if ea["is_affected"]:
                    affected_results.append(summary)
                else:
                    comparison_results.append(summary)

                # 콘솔 요약
                tag = "🎯 영향종목" if ea["is_affected"] else "🔄 비교종목"
                if ea.get("status") == "future_or_no_data_yet":
                    b5 = ea.get("before_5d", {})
                    print(f"    {tag} {sd['name']:14s} 사전 5일 평균 등락 {b5.get('avg_change_pct', 0):+.2f}% "
                          f"외인 누적 {b5.get('total_foreign_net', 0):+,}")
                    print(f"        ⚠️ {ea.get('message', '')}")
                elif ea.get("event_day"):
                    print(f"    {tag} {sd['name']:14s} 당일 등락 {ea['event_day']['change_pct']:+.2f}% "
                          f"거래량 {ea['event_day']['volume']:,} "
                          f"외인 {ea['event_day']['foreign_net_shares']:+,}")
                if ea.get("trend_changes"):
                    for tc in ea["trend_changes"]:
                        print(f"        → {tc}")

        event_reports.append({
            "label": evt["label"],
            "date": evt["date"],
            "description": evt["description"],
            "affected": affected_results,
            "comparison": comparison_results,
        })

    # 출력 저장
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "methodology": (
            "Naver 금융 frgn.naver 페이지 3개 (60거래일) 스크래핑 + "
            "자동 이벤트 감지 (거래량 +200% / 등락 ±3% / 외인 이상치) + "
            "사전 정의 이벤트 전후 통계 비교"
        ),
        "target_count": len(stocks_data),
        "pages_per_stock": PAGES_TO_FETCH,
        "days_per_stock": PAGES_TO_FETCH * 20,
        "predefined_events": PREDEFINED_EVENTS,
        "user_events": state.get("user_events", []),
        "stocks": stocks_data,
        "event_reports": event_reports,
        "data_source": "네이버 금융 (finance.naver.com)",
        "warning": "🚨 개인 매매는 추정값 (외인+기관 부호 반전). 자동 매매 금지.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    save_state(STATE_NAME, state)

    print("\n" + "=" * 72)
    print(f"  완료: {len(stocks_data)}개 종목 · {len(event_reports)}개 사전 정의 이벤트 분석")
    print("=" * 72)


if __name__ == "__main__":
    main()
