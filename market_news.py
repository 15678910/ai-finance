"""
시장 뉴스 수집기 (급변 원인 파악)
=================================
무료 RSS/페이지에서 미국·한국 시장 뉴스를 수집해 카테고리 분류.
- 미국: Yahoo Finance RSS (^GSPC, ^IXIC, ^SOX)
- 한국: 네이버 금융 뉴스 (증시)
키 불필요. 출력: docs/market_news.json
🚨 정보 모니터링용. 투자 결정 단독 사용 금지.
"""

import json
import os
import re
import sys
import html as _html
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "market_news.json")
UA = "Mozilla/5.0 (compatible; ai-finance-news/1.0)"

# 카테고리 키워드 (제목 매칭) — 우선순위 순
CATEGORIES = [
    ("지정학", "🌍", ["iran", "war", "geopolit", "middle east", "russia", "ukraine", "missile", "전쟁", "중동", "지정학", "이란", "북한"]),
    ("Fed·물가", "🏦", ["fed", "cpi", "inflation", "rate cut", "rate hike", "fomc", "powell", "ppi", "treasury", "yield", "금리", "물가", "인플레", "연준", "국채"]),
    ("AI·반도체", "🔌", ["nvidia", "chip stock", "semiconductor", "artificial intelligence", "tsmc", "micron", "broadcom", "데이터센터", "반도체", "엔비디아", "마이크론", "hbm", " a.i."]),
    ("실적", "📈", ["earnings", "revenue", "guidance", "results", "실적", "매출", "어닝"]),
    ("암호화폐", "🪙", ["bitcoin", "crypto", "ethereum", "비트코인", "코인", "이더리움"]),
    ("기타", "📰", []),
]


def _fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=12) as r:
        return r.read().decode("utf-8", errors="replace")


def _categorize(title):
    t = title.lower()
    for name, emoji, kws in CATEGORIES:
        if any(k in t for k in kws):
            return name, emoji
    return "기타", "📰"


def fetch_yahoo(symbol):
    """Yahoo Finance RSS → [{title, link, pub}]."""
    out = []
    try:
        url = f"https://feeds.finance.yahoo.com/rss/2.0/headline?s={symbol}&region=US&lang=en-US"
        txt = _fetch(url)
        items = re.findall(r"<item>(.*?)</item>", txt, re.S)
        for it in items:
            tm = re.search(r"<title>(.*?)</title>", it, re.S)
            lm = re.search(r"<link>(.*?)</link>", it, re.S)
            pm = re.search(r"<pubDate>(.*?)</pubDate>", it, re.S)
            if not tm:
                continue
            title = _html.unescape(re.sub(r"<.*?>", "", tm.group(1)).strip())
            if title.startswith("Yahoo"):
                continue
            out.append({"title": title,
                        "link": (lm.group(1).strip() if lm else ""),
                        "pub": (pm.group(1).strip() if pm else ""),
                        "region": "🇺🇸"})
    except Exception as e:
        print(f"  [WARN] Yahoo {symbol} 실패: {e}")
    return out


def fetch_naver():
    """네이버 금융 증시 뉴스 → [{title, link}]."""
    out = []
    try:
        url = "https://finance.naver.com/news/news_list.naver?mode=LSS2D&section_id=101&section_id2=258"
        txt = _fetch(url)
        # EUC-KR 페이지 → 재디코드
        raw = urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": UA}), timeout=12).read()
        try:
            txt = raw.decode("euc-kr", errors="replace")
        except Exception:
            txt = raw.decode("utf-8", errors="replace")
        subs = re.findall(r'class="articleSubject"[^>]*>\s*<a[^>]*href="([^"]+)"[^>]*>([^<]+)', txt)
        for href, title in subs[:15]:
            title = _html.unescape(title.strip())
            link = "https://finance.naver.com" + href if href.startswith("/") else href
            out.append({"title": title, "link": link, "pub": "", "region": "🇰🇷"})
    except Exception as e:
        print(f"  [WARN] 네이버 뉴스 실패: {e}")
    return out


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    print("=" * 55)
    print("  시장 뉴스 수집기")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    raw = []
    for sym in ("^GSPC", "^IXIC", "^SOX"):
        raw += fetch_yahoo(sym)
    raw += fetch_naver()

    # 중복 제거 + 카테고리
    from news_impact import classify_news, aggregate_sentiment

    seen = set()
    items = []
    for n in raw:
        key = n["title"][:60]
        if key in seen:
            continue
        seen.add(key)
        cat, emoji = _categorize(n["title"])
        sc, slabel, semoji, impact = classify_news(n["title"])
        items.append({**n, "category": cat, "cat_emoji": emoji,
                      "sentiment": sc, "sent_label": slabel, "sent_emoji": semoji, "impact": impact})

    # 카테고리 우선순위로 정렬 (지정학·Fed·반도체 먼저)
    order = {c[0]: i for i, c in enumerate(CATEGORIES)}
    items.sort(key=lambda x: order.get(x["category"], 99))
    items = items[:24]

    from collections import Counter
    cnt = Counter(i["category"] for i in items)
    # 지정학·Fed·시장 헤드라인 중심으로 순심리 산출
    sent_titles = [i["title"] for i in items if i["category"] in ("지정학", "Fed·물가", "AI·반도체")]
    sentiment = aggregate_sentiment(sent_titles)
    print(f"  수집 {len(items)}건 · 카테고리: {dict(cnt)}")
    print(f"  뉴스 심리: {sentiment['emoji']} {sentiment['label']} (완화 {sentiment['deesc']} · 격화 {sentiment['esc']} · 순 {sentiment['score']:+d})")

    # ── 예측 반전경보: 선물 예측 방향 ↔ 뉴스 심리 괴리 감지 ──
    reversal = None
    try:
        with open(os.path.join(BASE_DIR, "docs", "kospi_scenario.json"), encoding="utf-8") as f:
            fut = json.load(f).get("futures") or {}
        fpct = fut.get("predicted_pct")
        if fpct is not None and sentiment["score"] != 0:
            fdir = 1 if fpct > 0 else -1  # 선물 예측 방향
            sdir = 1 if sentiment["score"] > 0 else -1  # 뉴스 심리 방향
            if fdir < 0 and sdir > 0:
                reversal = {"flag": "up", "emoji": "🟢⚠️",
                            "text": f"선물은 하락 신호({fpct:+.1f}%)지만 뉴스는 완화(순{sentiment['score']:+d}) → 반등 주의",
                            "fut_pct": fpct}
            elif fdir > 0 and sdir < 0:
                reversal = {"flag": "down", "emoji": "🔴⚠️",
                            "text": f"선물은 상승 신호({fpct:+.1f}%)지만 뉴스는 격화(순{sentiment['score']:+d}) → 급락 반전 주의",
                            "fut_pct": fpct}
        if reversal:
            print(f"  ⚠️ 반전경보: {reversal['text']}")
    except Exception as e:
        print(f"  [WARN] 반전경보 계산 실패: {e}")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "items": items,
        "category_counts": dict(cnt),
        "sentiment": sentiment,
        "reversal_warning": reversal,
        "note": "Yahoo Finance·네이버 금융 무료 수집. 정보 모니터링용. 투자 결정 단독 사용 금지.",
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
