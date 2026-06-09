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
    seen = set()
    items = []
    for n in raw:
        key = n["title"][:60]
        if key in seen:
            continue
        seen.add(key)
        cat, emoji = _categorize(n["title"])
        items.append({**n, "category": cat, "cat_emoji": emoji})

    # 카테고리 우선순위로 정렬 (지정학·Fed·반도체 먼저)
    order = {c[0]: i for i, c in enumerate(CATEGORIES)}
    items.sort(key=lambda x: order.get(x["category"], 99))
    items = items[:24]

    from collections import Counter
    cnt = Counter(i["category"] for i in items)
    print(f"  수집 {len(items)}건 · 카테고리: {dict(cnt)}")
    for i in items[:6]:
        print(f"    [{i['category']}] {i['title'][:70]}")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "items": items,
        "category_counts": dict(cnt),
        "note": "Yahoo Finance·네이버 금융 무료 수집. 정보 모니터링용. 투자 결정 단독 사용 금지.",
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
