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
import urllib.parse
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "market_news.json")
UA = "Mozilla/5.0 (compatible; ai-finance-news/1.0)"

# 카테고리 키워드 (제목 매칭) — 우선순위 순
# 표기 규칙은 news_impact 와 동일: "war"=단어 단위, "geopolit*"=어간, 한글=부분 일치.
# (부분 일치로 두면 fed ⊂ FedEx, war ⊂ software 처럼 엉뚱한 기사가 끌려온다)
CATEGORIES = [
    ("지정학", "🌍", ["iran", "israel", "gaza", "war", "wars", "geopolit*", "middle east",
                     "russia*", "ukrain*", "missile*", "strike", "strikes", "sanction*", "opec",
                     "tariff*", "trade war", "ceasefire", "nato", "summit", "diplomac*",
                     "foreign policy", "vance", "pakistan", "north korea", "taiwan strait",
                     "전쟁", "중동", "지정학", "이란", "이스라엘", "가자", "북한", "우크라", "러시아",
                     "외교", "정상회담", "나토", "휴전", "관세", "제재", "파키스탄"]),
    ("Fed·물가", "🏦", ["fed", "feds", "federal reserve", "cpi", "inflation", "deflation",
                       "rate cut*", "rate hike*", "rate decision", "fomc", "powell", "bessent",
                       "ppi", "treasury", "treasuries", "yield", "yields", "bond market",
                       "basis point*", "jobs report", "payroll*", "jackson hole",
                       "금리", "물가", "인플레", "연준", "국채", "베센트", "파월", "고용지표"]),
    ("AI·반도체", "🔌", ["nvidia", "chip stock*", "semiconductor*", "artificial intelligence",
                        "tsmc", "micron", "broadcom", "데이터센터", "반도체", "엔비디아",
                        "마이크론", "hbm", "a.i."]),
    ("실적", "📈", ["earnings", "revenue", "guidance", "results", "실적", "매출", "어닝"]),
    ("암호화폐", "🪙", ["bitcoin", "btc", "crypto*", "ethereum", "stablecoin*", "digital asset*",
                      "spot etf", "sec approval", "비트코인", "코인", "이더리움", "가상자산",
                      "스테이블코인"]),
    ("기타", "📰", []),
]

# 시장과 무관한데 키워드만 걸리는 기사 — 삭제하지 않고 우선순위만 낮춘다.
# 실측 사례: 'Women's Global Sports Summit'(summit), 'lightning strike ignites
# attic fire in Summit, NE'(strike+summit) 가 지정학 '하락 위험'으로 분류됐다.
NOISE = ["sports", "tournament", "championship", "playoff*", "nfl", "nba", "mlb", "nhl",
         "soccer", "olympic*", "athlete*", "lightning strike", "attic", "wildfire",
         "high school", "obituary", "funeral", "arrested", "shooting", "weather forecast",
         "recipe", "celebrity", "box office"]


def _fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=12) as r:
        return r.read().decode("utf-8", errors="replace")


from news_impact import _compile as _compile_kw   # 단어 경계 매칭 (표기 규칙 공유)

_CAT_RE = [(name, emoji, _compile_kw(kws) if kws else None) for name, emoji, kws in CATEGORIES]
_NOISE_RE = _compile_kw(NOISE)


def _categorize(title):
    t = title.lower()
    for name, emoji, rx in _CAT_RE:
        if rx and rx.search(t):
            return name, emoji
    return "기타", "📰"


def is_noise(title):
    """시장과 무관한 기사인지 — 키워드만 우연히 걸린 지역·스포츠·사건사고 기사."""
    return bool(_NOISE_RE.search((title or "").lower()))


# 중요도 가중 — 24건 상한과 브리핑 톱3 선별에 쓴다.
# 이전에는 최신순으로만 잘라서, 시장을 움직인 헤드라인이 몇 시간 늦었다는
# 이유만으로 투자 칼럼("If a Bear Market Is Coming…")에 밀려 잘려나갔다.
CAT_WEIGHT = {"Fed·물가": 3, "지정학": 2, "AI·반도체": 2, "암호화폐": 2, "실적": 1, "기타": 0}


def priority_score(category, sentiment, noise=False):
    """기사 중요도 (높을수록 우선). 노이즈는 음수로 밀어낸다."""
    if noise:
        return -5
    score = CAT_WEIGHT.get(category, 0)
    if sentiment != 0:          # 격화/완화 판정이 붙은 = 방향성 있는 기사
        score += 3
    return score


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


def fetch_rss(url, region="🌍", limit=10):
    """범용 RSS(item: title/link/pubDate) → [{title, link, pub, region}]."""
    out = []
    try:
        txt = _fetch(url)
        items = re.findall(r"<item>(.*?)</item>", txt, re.S)
        for it in items[:limit]:
            tm = re.search(r"<title>(.*?)</title>", it, re.S)
            lm = re.search(r"<link>(.*?)</link>", it, re.S)
            pm = re.search(r"<pubDate>(.*?)</pubDate>", it, re.S)
            if not tm:
                continue
            title = _html.unescape(re.sub(r"<.*?>", "", tm.group(1)).strip())
            if not title:
                continue
            out.append({"title": title,
                        "link": (lm.group(1).strip() if lm else ""),
                        "pub": (pm.group(1).strip() if pm else ""),
                        "region": region})
    except Exception as e:
        print(f"  [WARN] RSS 실패 ({url[:40]}…): {e}")
    return out


def _gnews(q):
    """Google News RSS 검색 — 키 불필요·최근순. 정치/외교/지정학 뉴스 보강용."""
    return ("https://news.google.com/rss/search?q=" + urllib.parse.quote(q)
            + "&hl=en-US&gl=US&ceid=US:en")


# 주제별 보강 피드 — Yahoo 지수 RSS(^GSPC 등)는 종목·증시 칼럼 위주라
# 정책·통화·규제 헤드라인을 거의 싣지 않는다. 카테고리 분류기에는 'Fed·물가'와
# '암호화폐'가 있는데 정작 그 뉴스를 가져오는 소스가 없어 항상 0~1건이었다.
# (2026-08 실측: 24건 중 암호화폐 0건 · Fed·물가 1건)
TOPIC_FEEDS = [
    # 지정학·외교 (기존)
    (_gnews('(Iran OR Israel OR "Middle East" OR sanctions OR OPEC OR ceasefire) when:3d'), "🌍"),
    (_gnews('("foreign policy" OR summit OR diplomacy OR NATO OR tariff OR "trade war") when:3d'), "🌍"),
    # 연준·국채·통화정책 — 베센트 재무장관·파월 의장 발언이 여기로 들어온다
    (_gnews('(Bessent OR Powell OR "Federal Reserve" OR FOMC OR "Treasury yield" '
            'OR "bond market" OR "rate cut" OR "rate hike") when:2d'), "🇺🇸"),
    (_gnews('(CPI OR inflation OR "jobs report" OR payrolls OR "Jackson Hole") when:2d'), "🇺🇸"),
    # 암호화폐 — 규제·정책 발언 포함
    (_gnews('(bitcoin OR crypto OR stablecoin OR "digital assets" OR "SEC crypto") when:2d'), "🌐"),
]


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
    for url, region in TOPIC_FEEDS:   # 지정학·연준/국채·암호화폐 보강
        raw += fetch_rss(url, region)

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
        noise = is_noise(n["title"])
        items.append({**n, "category": cat, "cat_emoji": emoji,
                      "sentiment": sc, "sent_label": slabel, "sent_emoji": semoji,
                      "impact": impact, "noise": noise,
                      "priority": priority_score(cat, sc, noise)})

    # 최근순 정렬 + 오래된 기사(>4일) 제거 — 카테고리순이면 묵은 기사가 위로 올라와 '정체'처럼 보임
    from email.utils import parsedate_to_datetime
    from datetime import timezone as _tz, timedelta as _td
    _now = datetime.now(_tz.utc)

    def _pub(it):
        try:
            return parsedate_to_datetime(it.get("pub", "")).astimezone(_tz.utc)
        except Exception:
            return None
    items = [it for it in items if (_pub(it) is None) or (_now - _pub(it)) <= _td(days=4)]

    # 24건 상한은 '중요도' 로 자르고, 살아남은 것을 '최신순' 으로 보여준다.
    # 최신순으로 자르면 몇 시간 늦은 정책 헤드라인이 방금 올라온 투자 칼럼에 밀린다.
    _oldest = _now - _td(days=999)
    items.sort(key=lambda it: (it["priority"], _pub(it) or _oldest), reverse=True)
    items = items[:24]
    items.sort(key=lambda it: (_pub(it) or _oldest), reverse=True)   # 표시는 최신 먼저

    from collections import Counter
    cnt = Counter(i["category"] for i in items)
    # 지정학·Fed·시장 헤드라인 중심으로 순심리 산출 (노이즈는 제외 —
    # '낙뢰 strike' 같은 기사가 격화로 잡히면 반전 경보까지 오염된다)
    sent_titles = [i["title"] for i in items
                   if i["category"] in ("지정학", "Fed·물가", "AI·반도체") and not i["noise"]]
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
