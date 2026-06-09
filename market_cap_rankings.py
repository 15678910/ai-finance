"""
시가총액 순위 수집기 (국내·글로벌)
====================================
- 국내: KOSPI 상위 종목 (yfinance .KS)
- 글로벌: 세계 시총 상위 종목 (yfinance)

출력: docs/market_cap_rankings.json
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "market_cap_rankings.json")

# ── 국내 추적 종목 (코스피·코스닥 시총 상위) ─────────────────────────────
DOMESTIC_TICKERS = [
    ("005930", "삼성전자",     "🇰🇷", "IT/반도체"),
    ("000660", "SK하이닉스",   "🇰🇷", "IT/반도체"),
    ("005380", "현대차",       "🇰🇷", "자동차"),
    ("373220", "LG에너지솔루션","🇰🇷","배터리"),
    ("207940", "삼성바이오로직스","🇰🇷","바이오"),
    ("000270", "기아",         "🇰🇷", "자동차"),
    ("105560", "KB금융",       "🇰🇷", "금융"),
    ("055550", "신한지주",     "🇰🇷", "금융"),
    ("006400", "삼성SDI",      "🇰🇷", "배터리"),
    ("068270", "셀트리온",     "🇰🇷", "바이오"),
    ("012450", "한화에어로스페이스","🇰🇷","방산"),
    ("051910", "LG화학",       "🇰🇷", "화학"),
    ("012330", "현대모비스",   "🇰🇷", "자동차"),
    ("035420", "네이버",       "🇰🇷", "IT플랫폼"),
    ("086790", "하나금융",     "🇰🇷", "금융"),
    ("066570", "LG전자",       "🇰🇷", "전자"),
    ("316140", "우리금융",     "🇰🇷", "금융"),
    ("005490", "POSCO홀딩스",  "🇰🇷", "철강/소재"),
    ("035720", "카카오",       "🇰🇷", "IT플랫폼"),
    ("138040", "메리츠금융",   "🇰🇷", "금융"),
    ("267250", "HD현대",       "🇰🇷", "조선"),
    ("128940", "한미약품",     "🇰🇷", "바이오"),
    ("079550", "LIG넥스원",   "🇰🇷", "방산"),
    ("247540", "에코프로비엠", "🇰🇷", "배터리소재"),
    ("009540", "HD한국조선해양","🇰🇷","조선"),
]

# ── 글로벌 추적 종목 (세계 시총 상위) ──────────────────────────────────────
GLOBAL_TICKERS = [
    ("AAPL",    "애플",          "🇺🇸", "IT/소비재"),
    ("MSFT",    "마이크로소프트", "🇺🇸", "IT/플랫폼"),
    ("NVDA",    "엔비디아",      "🇺🇸", "반도체/AI"),
    ("GOOGL",   "알파벳",        "🇺🇸", "IT/플랫폼"),
    ("AMZN",    "아마존",        "🇺🇸", "IT/유통"),
    ("META",    "메타",          "🇺🇸", "IT/플랫폼"),
    ("TSLA",    "테슬라",        "🇺🇸", "전기차"),
    ("AVGO",    "브로드컴",      "🇺🇸", "반도체"),
    ("BRK-B",   "버크셔해서웨이","🇺🇸", "금융"),
    ("TSM",     "TSMC",          "🇹🇼", "반도체"),
    ("LLY",     "일라이릴리",    "🇺🇸", "바이오"),
    ("JPM",     "JP모건",        "🇺🇸", "금융"),
    ("V",       "비자",          "🇺🇸", "금융"),
    ("UNH",     "유나이티드헬스", "🇺🇸", "헬스케어"),
    ("XOM",     "엑슨모빌",      "🇺🇸", "에너지"),
    ("WMT",     "월마트",        "🇺🇸", "유통"),
    ("ASML",    "ASML",          "🇳🇱", "반도체장비"),
    ("MA",      "마스터카드",    "🇺🇸", "금융"),
    ("COST",    "코스트코",      "🇺🇸", "유통"),
    ("005930.KS","삼성전자",     "🇰🇷", "IT/반도체"),
    ("000660.KS","SK하이닉스",  "🇰🇷", "IT/반도체"),
    ("NVO",     "노보노디스크",  "🇩🇰", "바이오"),
    ("ORCL",    "오라클",        "🇺🇸", "IT/클라우드"),
    ("AMD",     "AMD",           "🇺🇸", "반도체"),
    ("MU",      "마이크론",      "🇺🇸", "반도체/메모리"),
    ("INTC",    "인텔",          "🇺🇸", "반도체"),
    ("QCOM",    "퀄컴",          "🇺🇸", "반도체"),
    ("SAP",     "SAP",           "🇩🇪", "IT/소프트웨어"),
]


def _naver_kr(ticker: str):
    """네이버 실시간으로 한국주 (현재가, 등락률%) — yfinance .KS 장중 stale 회피."""
    import urllib.request
    try:
        url = f"https://polling.finance.naver.com/api/realtime/domestic/stock/{ticker}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        d = json.loads(urllib.request.urlopen(req, timeout=10).read())["datas"][0]
        price = float(str(d.get("closePrice", "")).replace(",", ""))
        chg = float(str(d.get("fluctuationsRatio", "")).replace(",", ""))
        return price, chg
    except Exception:
        return None, None


def fetch_stock(ticker_id: str, name: str, flag: str, sector: str, is_krw: bool = False) -> dict | None:
    """yfinance로 단일 종목 시총·가격 수집 (한국주는 네이버로 가격·등락 보정)."""
    try:
        import yfinance as yf
        yt = ticker_id + ".KS" if (is_krw and not ticker_id.endswith(".KS")) else ticker_id
        t = yf.Ticker(yt)
        fi = t.fast_info
        price = getattr(fi, "last_price", None)
        prev  = getattr(fi, "previous_close", None)
        mc    = getattr(fi, "market_cap", None)
        if not mc:
            info = t.info or {}
            mc = info.get("marketCap")
            if not price:
                price = info.get("currentPrice") or info.get("regularMarketPrice")
            if not prev:
                prev = info.get("regularMarketPreviousClose") or info.get("previousClose")

        if not mc or not price:
            return None

        mc_f = float(mc)
        chg = round((float(price) - float(prev)) / float(prev) * 100, 2) if prev and float(prev) > 0 else 0.0

        # 한국주: 네이버로 가격·등락 보정 (yfinance .KS 장중 stale → 부호 반대 회피)
        if is_krw:
            n_price, n_chg = _naver_kr(ticker_id)
            if n_price and float(price) > 0:
                mc_f = mc_f * (n_price / float(price))  # 시총 비례 보정
                price = n_price
            if n_chg is not None:
                chg = round(n_chg, 2)

        # 국내: 원화 → 조원, 해외: 달러 → 조달러
        mc_display = round(mc_f / 1e12, 2) if is_krw else round(mc_f / 1e12, 3)

        return {
            "ticker":   ticker_id,
            "name":     name,
            "flag":     flag,
            "sector":   sector,
            "price":    round(float(price), 0 if is_krw else 2),
            "market_cap": mc_display,
            "change_pct": chg,
            "unit":     "조원" if is_krw else "조달러",
        }
    except Exception as e:
        print(f"  [SKIP] {name} ({ticker_id}): {e}")
        return None


def main():
    print("=" * 55)
    print("  시가총액 순위 수집")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    # 국내 순위
    print("\n[국내] 수집 중...")
    domestic = []
    for tid, name, flag, sector in DOMESTIC_TICKERS:
        r = fetch_stock(tid, name, flag, sector, is_krw=True)
        if r:
            domestic.append(r)
            print(f"  {name}: {r['market_cap']}조원 ({r['change_pct']:+.2f}%)")
    domestic.sort(key=lambda x: x["market_cap"], reverse=True)

    # 글로벌 순위
    print("\n[글로벌] 수집 중...")
    global_stocks = []
    seen = set()
    for tid, name, flag, sector in GLOBAL_TICKERS:
        r = fetch_stock(tid, name, flag, sector, is_krw=(".KS" in tid))
        if r and name not in seen:
            seen.add(name)
            # 글로벌에서 한국주는 조원 → 조달러 환산 (÷1500)
            if ".KS" in tid and r["unit"] == "조원":
                r["market_cap"] = round(r["market_cap"] / 1500, 3)
                r["unit"] = "조달러"
            global_stocks.append(r)
            print(f"  {name}: {r['market_cap']}조달러 ({r['change_pct']:+.2f}%)")
    global_stocks.sort(key=lambda x: x["market_cap"], reverse=True)

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "domestic": domestic[:20],
        "global":   global_stocks[:20],
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    print(f"  국내 {len(domestic)}개 / 글로벌 {len(global_stocks)}개")
    return 0


if __name__ == "__main__":
    sys.exit(main())
