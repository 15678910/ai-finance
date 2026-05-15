"""
한국 시장 투자자별 매매 동향 모니터 (외국인·기관)
=================================================

목적:
  KOSPI/KOSDAQ 주요 종목들의 외국인·기관 순매매 동향을 매일 추적.
  강력한 매수/매도 시그널 자동 감지 → 텔레그램 알림 + 대시보드.

데이터 소스: 네이버 금융 종목별 외국인/기관 매매 페이지
  · https://finance.naver.com/item/frgn.naver?code={ticker}
  · 무료, API 키 불필요
  · 일별 데이터 (장 마감 후 갱신)
  · 컬럼: 날짜, 종가, 등락률, 거래량, 기관 순매매(주식수), 외국인 순매매(주식수),
           외국인 보유 주식수, 외국인 보유율

추적 시그널 (강도 ⭐):
  ⭐⭐⭐⭐⭐ 외인+기관 동시 5일+ 연속 매수 (강력한 합의 매수)
  ⭐⭐⭐⭐⭐ 외인+기관 동시 5일+ 연속 매도 (강력한 합의 매도)
  ⭐⭐⭐⭐  외인 5일+ 연속 매수 또는 매도
  ⭐⭐⭐⭐  외인-기관 의견 충돌 (한쪽 매수, 한쪽 매도)
  ⭐⭐⭐   외인 보유율 5일간 1%p+ 변동
  ⭐⭐⭐   거래대금 대비 큰 비중 매매

알림: 매일 KST 16:00 (장 마감 후 30분), 강력 시그널 즉시 텔레그램

🚨 시뮬레이션. 자동 매매 금지.
"""

import os
import sys
import re
import json
import time
import urllib.request
import urllib.error
from datetime import datetime, timezone, timedelta

from core import send_message, load_state, save_state

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "investor_flow.json")
STATE_NAME = "investor_flow"
KST = timezone(timedelta(hours=9))

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
TIMEOUT_SEC = 15

# 추적 대상 종목 (value_screener의 universe + 핵심 종목 추가)
TARGET_STOCKS = [
    # IT/반도체
    {"ticker": "005930", "name": "삼성전자", "sector": "IT/반도체"},
    {"ticker": "000660", "name": "SK하이닉스", "sector": "IT/반도체"},
    {"ticker": "042700", "name": "한미반도체", "sector": "반도체장비"},
    {"ticker": "000990", "name": "DB하이텍", "sector": "반도체"},
    {"ticker": "035420", "name": "네이버", "sector": "IT/플랫폼"},
    {"ticker": "035720", "name": "카카오", "sector": "IT/플랫폼"},
    {"ticker": "066570", "name": "LG전자", "sector": "전자"},
    {"ticker": "018260", "name": "삼성에스디에스", "sector": "IT"},
    {"ticker": "034020", "name": "두산에너빌리티", "sector": "IT/원전"},
    # 금융
    {"ticker": "055550", "name": "신한지주", "sector": "금융"},
    {"ticker": "086790", "name": "하나금융지주", "sector": "금융"},
    {"ticker": "105560", "name": "KB금융", "sector": "금융"},
    {"ticker": "316140", "name": "우리금융지주", "sector": "금융"},
    {"ticker": "024110", "name": "기업은행", "sector": "금융"},
    # 자동차
    {"ticker": "005380", "name": "현대차", "sector": "자동차"},
    {"ticker": "000270", "name": "기아", "sector": "자동차"},
    {"ticker": "012330", "name": "현대모비스", "sector": "자동차"},
    # 화학/에너지
    {"ticker": "051910", "name": "LG화학", "sector": "화학"},
    {"ticker": "010950", "name": "S-Oil", "sector": "에너지"},
    {"ticker": "096770", "name": "SK이노베이션", "sector": "에너지"},
    {"ticker": "015760", "name": "한국전력", "sector": "유틸리티"},
    # 철강/조선/방산
    {"ticker": "005490", "name": "POSCO홀딩스", "sector": "철강"},
    {"ticker": "009540", "name": "HD한국조선해양", "sector": "조선"},
    {"ticker": "012450", "name": "한화에어로스페이스", "sector": "방산"},
    {"ticker": "079550", "name": "LIG넥스원", "sector": "방산"},
    {"ticker": "064350", "name": "현대로템", "sector": "방산"},
    # 바이오/제약
    {"ticker": "068270", "name": "셀트리온", "sector": "바이오"},
    {"ticker": "207940", "name": "삼성바이오로직스", "sector": "바이오"},
    {"ticker": "128940", "name": "한미약품", "sector": "제약"},
    {"ticker": "000100", "name": "유한양행", "sector": "제약"},
    # 배터리
    {"ticker": "373220", "name": "LG에너지솔루션", "sector": "배터리"},
    {"ticker": "006400", "name": "삼성SDI", "sector": "배터리"},
    # 통신
    {"ticker": "017670", "name": "SK텔레콤", "sector": "통신"},
    {"ticker": "030200", "name": "KT", "sector": "통신"},
    # 소비재·유통
    {"ticker": "271560", "name": "오리온", "sector": "소비재"},
    {"ticker": "097950", "name": "CJ제일제당", "sector": "소비재"},
    {"ticker": "033780", "name": "KT&G", "sector": "고배당"},
    {"ticker": "139480", "name": "이마트", "sector": "유통"},
    # 게임/엔터
    {"ticker": "036570", "name": "엔씨소프트", "sector": "게임"},
    {"ticker": "259960", "name": "크래프톤", "sector": "게임"},
    {"ticker": "352820", "name": "하이브", "sector": "엔터"},
    # 항공·화장품
    {"ticker": "003490", "name": "대한항공", "sector": "항공"},
    {"ticker": "090430", "name": "아모레퍼시픽", "sector": "화장품"},
    {"ticker": "051900", "name": "LG생활건강", "sector": "화장품"},
    # 건설/물류
    {"ticker": "000720", "name": "현대건설", "sector": "건설"},
    {"ticker": "000120", "name": "CJ대한통운", "sector": "물류"},
]


# ====================================================================
# 네이버 금융 스크래핑
# ====================================================================
def fetch_naver_flow(ticker: str) -> list:
    """종목별 외국인·기관 매매 동향 (최근 ~20일).

    Returns: [{date, close, change_pct, volume,
               institutional_net_shares, foreign_net_shares,
               foreign_holding_shares, foreign_holding_pct}]
    """
    url = f"https://finance.naver.com/item/frgn.naver?code={ticker}"
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": USER_AGENT,
            "Referer": "https://finance.naver.com/",
        })
        with urllib.request.urlopen(req, timeout=TIMEOUT_SEC) as resp:  # nosec — 공개 페이지
            html = resp.read().decode("euc-kr", errors="replace")
    except Exception as e:
        print(f"    [실패] {ticker}: {e}")
        return []

    # 데이터 행 추출 (날짜 형식 포함하는 tr만)
    tr_pattern = re.compile(r"<tr[^>]*>(.*?)</tr>", re.DOTALL)
    date_pattern = re.compile(r"(\d{4}\.\d{2}\.\d{2})")
    rows = []

    for tr in tr_pattern.findall(html):
        date_match = date_pattern.search(tr)
        if not date_match:
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
            rows.append(row)
        except Exception:
            continue
    return rows


def _parse_int(s: str) -> int:
    """+260,989 / -1,234 / -- → int"""
    if not s or s in ("-", "--", "N/A"):
        return 0
    s = s.replace(",", "").replace("+", "").strip()
    try:
        return int(float(s))
    except Exception:
        return 0


def _parse_pct(s: str) -> float:
    """+4.23% / -1.10% → float (소수점, 비율)"""
    if not s or s in ("-", "--"):
        return 0.0
    s = s.replace("%", "").replace("+", "").replace(",", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0


# ====================================================================
# 시그널 감지
# ====================================================================
def detect_signals(stock_info: dict, rows: list) -> dict:
    """투자자 매매 동향 분석 + 시그널 감지."""
    if not rows or len(rows) < 2:
        return {"signals": [], "summary": "데이터 부족"}

    # 최신순으로 정렬 보장 (날짜 내림차순)
    rows_sorted = sorted(rows, key=lambda r: r["date"], reverse=True)
    recent = rows_sorted[:5]   # 최근 5거래일
    week = rows_sorted[:5]
    twoweek = rows_sorted[:10] if len(rows_sorted) >= 10 else rows_sorted

    # 연속 일수 계산
    def streak(rows_list, key):
        """최신 행부터 같은 부호(>0 or <0)가 연속된 일수."""
        if not rows_list:
            return 0, 0  # streak, direction
        first_val = rows_list[0][key]
        if first_val == 0:
            return 0, 0
        direction = 1 if first_val > 0 else -1
        s = 0
        for r in rows_list:
            v = r[key]
            if (v > 0 and direction == 1) or (v < 0 and direction == -1):
                s += 1
            else:
                break
        return s, direction

    foreign_streak, foreign_dir = streak(rows_sorted, "foreign_net_shares")
    inst_streak, inst_dir = streak(rows_sorted, "institutional_net_shares")

    # 5일 누적 순매수
    foreign_5d = sum(r["foreign_net_shares"] for r in week)
    inst_5d = sum(r["institutional_net_shares"] for r in week)
    foreign_10d = sum(r["foreign_net_shares"] for r in twoweek)

    # 외국인 보유율 변화
    if len(twoweek) >= 5:
        holding_change = twoweek[0]["foreign_holding_pct"] - twoweek[-1]["foreign_holding_pct"]
    else:
        holding_change = 0.0

    # 시그널 감지
    signals = []

    # 1. 외인 + 기관 동시 강매수 (5일 모두 양수)
    if all(r["foreign_net_shares"] > 0 for r in week) and all(r["institutional_net_shares"] > 0 for r in week):
        signals.append({
            "type": "consensus_buy",
            "strength": 5,
            "emoji": "🚀",
            "label": "외인+기관 5일 동시 매수",
            "description": f"외인 {foreign_5d:+,} · 기관 {inst_5d:+,} (5일 누적)",
        })
    # 2. 외인 + 기관 동시 강매도 (5일 모두 음수)
    elif all(r["foreign_net_shares"] < 0 for r in week) and all(r["institutional_net_shares"] < 0 for r in week):
        signals.append({
            "type": "consensus_sell",
            "strength": 5,
            "emoji": "📉",
            "label": "외인+기관 5일 동시 매도",
            "description": f"외인 {foreign_5d:+,} · 기관 {inst_5d:+,} (5일 누적)",
        })
    # 3. 외인 5일+ 연속 (한 방향)
    elif foreign_streak >= 5:
        signals.append({
            "type": "foreign_streak_buy" if foreign_dir > 0 else "foreign_streak_sell",
            "strength": 4,
            "emoji": "🔥" if foreign_dir > 0 else "🩸",
            "label": f"외인 {foreign_streak}일 연속 {'매수' if foreign_dir > 0 else '매도'}",
            "description": f"누적 {foreign_5d:+,}주",
        })

    # 4. 외인-기관 의견 충돌
    if foreign_5d > 0 and inst_5d < 0 and abs(foreign_5d) > 10000 and abs(inst_5d) > 10000:
        signals.append({
            "type": "disagreement",
            "strength": 3,
            "emoji": "⚔️",
            "label": "외인 매수 vs 기관 매도",
            "description": f"외인 {foreign_5d:+,} ↑ · 기관 {inst_5d:+,} ↓",
        })
    elif foreign_5d < 0 and inst_5d > 0 and abs(foreign_5d) > 10000 and abs(inst_5d) > 10000:
        signals.append({
            "type": "disagreement",
            "strength": 3,
            "emoji": "⚔️",
            "label": "외인 매도 vs 기관 매수",
            "description": f"외인 {foreign_5d:+,} ↓ · 기관 {inst_5d:+,} ↑",
        })

    # 5. 외국인 보유율 큰 변동
    if abs(holding_change) >= 1.0:
        signals.append({
            "type": "holding_change",
            "strength": 3,
            "emoji": "📊",
            "label": f"외인 보유율 {holding_change:+.2f}%p 변동 (5일)",
            "description": f"현재 {twoweek[0]['foreign_holding_pct']:.2f}%",
        })

    # 종합 등급 (최고 강도 시그널)
    max_strength = max((s["strength"] for s in signals), default=0)
    if max_strength >= 5:
        rating, color = "강력 신호", "red"
    elif max_strength >= 4:
        rating, color = "주의 신호", "amber"
    elif max_strength >= 3:
        rating, color = "관찰", "cyan"
    else:
        rating, color = "평균", "gray"

    return {
        "signals": signals,
        "max_strength": max_strength,
        "rating": rating,
        "rating_color": color,
        "foreign_streak_days": foreign_streak,
        "foreign_streak_direction": "buy" if foreign_dir > 0 else ("sell" if foreign_dir < 0 else "none"),
        "institutional_streak_days": inst_streak,
        "institutional_streak_direction": "buy" if inst_dir > 0 else ("sell" if inst_dir < 0 else "none"),
        "foreign_5d_net_shares": foreign_5d,
        "institutional_5d_net_shares": inst_5d,
        "foreign_10d_net_shares": foreign_10d,
        "foreign_holding_pct_now": rows_sorted[0]["foreign_holding_pct"],
        "foreign_holding_pct_5d_change": round(holding_change, 3),
        "latest_close": rows_sorted[0]["close"],
        "latest_change_pct": rows_sorted[0]["change_pct"],
        "latest_date": rows_sorted[0]["date"],
    }


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 72)
    print("  한국 시장 투자자별 매매 동향 모니터 (외국인·기관)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  대상: {len(TARGET_STOCKS)}개 종목 (네이버 금융 스크래핑)")
    print("=" * 72)

    state = load_state(STATE_NAME, default={"last_alerted_signals": {}})

    # 종목별 분석
    print("\n[종목별 데이터 수집]")
    results = []
    for stock_info in TARGET_STOCKS:
        ticker = stock_info["ticker"]
        print(f"  {stock_info['name']:14s} ({ticker})...", end=" ")
        rows = fetch_naver_flow(ticker)
        if not rows:
            print("데이터 없음")
            continue
        analysis = detect_signals(stock_info, rows)

        results.append({
            "ticker": ticker,
            "name": stock_info["name"],
            "sector": stock_info.get("sector"),
            "rows": rows[:10],  # 최근 10일만 (JSON 크기 절감)
            **analysis,
        })

        sig_emoji = " ".join(s["emoji"] for s in analysis["signals"]) or "—"
        streak_info = f"외인 {analysis['foreign_streak_days']}일{analysis['foreign_streak_direction'][0]}"
        print(f"{sig_emoji} {streak_info} 보유율 {analysis['foreign_holding_pct_now']:.1f}%")
        time.sleep(0.25)  # rate limit (네이버 부담 방지)

    # 종합 정렬
    results.sort(key=lambda r: r.get("max_strength", 0), reverse=True)

    # 시그널별 Top 종목
    strong_signals = [r for r in results if r.get("max_strength", 0) >= 4]
    consensus_buy = [r for r in results if any(s["type"] == "consensus_buy" for s in r.get("signals", []))]
    consensus_sell = [r for r in results if any(s["type"] == "consensus_sell" for s in r.get("signals", []))]
    foreign_buy_streak = [r for r in results if any(s["type"] == "foreign_streak_buy" for s in r.get("signals", []))]
    foreign_sell_streak = [r for r in results if any(s["type"] == "foreign_streak_sell" for s in r.get("signals", []))]
    disagreement = [r for r in results if any(s["type"] == "disagreement" for s in r.get("signals", []))]

    # 외국인 보유율 변동 Top
    holding_increase = sorted(results, key=lambda r: r.get("foreign_holding_pct_5d_change", 0), reverse=True)[:10]
    holding_decrease = sorted(results, key=lambda r: r.get("foreign_holding_pct_5d_change", 0))[:10]

    # 텔레그램 알림 (강력 시그널 + 이전 알림과 다른 종목만)
    last_alerted = state.get("last_alerted_signals", {})
    new_alerts = []
    for r in strong_signals:
        sig_key = ":".join(sorted(s["type"] for s in r["signals"]))
        if last_alerted.get(r["ticker"]) != sig_key:
            new_alerts.append(r)
            last_alerted[r["ticker"]] = sig_key

    if new_alerts:
        lines = ["💰 한국 시장 투자자 매매 동향 — 강력 시그널", "=" * 30, ""]
        for r in new_alerts[:10]:
            primary_sig = r["signals"][0] if r["signals"] else None
            lines.append(f"{primary_sig['emoji'] if primary_sig else '•'} {r['name']} ({r['ticker']}) — {r['sector']}")
            for s in r["signals"][:2]:
                lines.append(f"   {s['label']}: {s['description']}")
            lines.append(f"   외인 보유율 {r['foreign_holding_pct_now']:.2f}% "
                         f"({r['foreign_holding_pct_5d_change']:+.2f}%p 5일)")
            lines.append("")
        lines.append("🚨 시뮬레이션. 자동 매매 금지.")
        lines.append("대시보드: https://15678910.github.io/ai-finance/")
        try:
            send_message("\n".join(lines))
            print(f"\n  ✅ 텔레그램 발송: 강력 시그널 {len(new_alerts)}건")
        except Exception as e:
            print(f"\n  ❌ 텔레그램 발송 실패: {e}")

    state["last_alerted_signals"] = last_alerted
    save_state(STATE_NAME, state)

    # 결과 저장
    trading_date = results[0]["latest_date"] if results else None
    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "trading_date": trading_date,
        "methodology": "네이버 금융 종목별 외국인·기관 매매 동향 스크래핑 + 시그널 자동 감지",
        "universe_size": len(TARGET_STOCKS),
        "analyzed_count": len(results),
        "strong_signals_count": len(strong_signals),
        "new_alerts_this_run": len(new_alerts),
        "results": results,
        "top_signals": {
            "consensus_buy": [_signal_summary(r) for r in consensus_buy],
            "consensus_sell": [_signal_summary(r) for r in consensus_sell],
            "foreign_buy_streak": [_signal_summary(r) for r in foreign_buy_streak],
            "foreign_sell_streak": [_signal_summary(r) for r in foreign_sell_streak],
            "disagreement": [_signal_summary(r) for r in disagreement],
            "foreign_holding_increase_top": [_signal_summary(r) for r in holding_increase if r.get("foreign_holding_pct_5d_change", 0) > 0][:5],
            "foreign_holding_decrease_top": [_signal_summary(r) for r in holding_decrease if r.get("foreign_holding_pct_5d_change", 0) < 0][:5],
        },
        "data_source": "네이버 금융 (finance.naver.com)",
        "warning": "🚨 일별 누적 데이터. 장중 실시간 아님. 시뮬레이션 전용.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    # 콘솔 요약
    print(f"\n[강력 시그널 Top 5]")
    for r in strong_signals[:5]:
        sig_label = r["signals"][0]["label"] if r["signals"] else "—"
        print(f"  {r['name']:14s} {r['signals'][0]['emoji'] if r['signals'] else ''} {sig_label}")

    print("\n" + "=" * 72)
    print(f"  완료: {len(results)}/{len(TARGET_STOCKS)} 종목 · 강력 시그널 {len(strong_signals)}건 · 신규 알림 {len(new_alerts)}건")
    print("=" * 72)


def _signal_summary(r: dict) -> dict:
    """JSON 출력용 종목 요약."""
    return {
        "ticker": r["ticker"],
        "name": r["name"],
        "sector": r.get("sector"),
        "foreign_5d_net_shares": r.get("foreign_5d_net_shares"),
        "institutional_5d_net_shares": r.get("institutional_5d_net_shares"),
        "foreign_holding_pct_now": r.get("foreign_holding_pct_now"),
        "foreign_holding_pct_5d_change": r.get("foreign_holding_pct_5d_change"),
        "foreign_streak_days": r.get("foreign_streak_days"),
        "foreign_streak_direction": r.get("foreign_streak_direction"),
        "rating": r.get("rating"),
        "rating_color": r.get("rating_color"),
        "primary_signal": (r["signals"][0]["label"] if r.get("signals") else None),
        "primary_signal_emoji": (r["signals"][0]["emoji"] if r.get("signals") else None),
    }


if __name__ == "__main__":
    main()
