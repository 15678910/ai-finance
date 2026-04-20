"""
긴급 뉴스 모니터링 시스템
========================
주요 언론사 RSS 피드를 체크하여 긴급 키워드가 포함된 뉴스를
텔레그램으로 즉시 전송합니다.

무료 (API 키 불필요), GitHub Actions 1시간 주기 실행.
"""

import os
import sys
import json
import re
import urllib.request
import urllib.parse
import urllib.error
import xml.etree.ElementTree as ET
from datetime import datetime, timezone, timedelta
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
STATE_FILE = os.path.join(BASE_DIR, "docs", "breaking_news_state.json")

# ====================================================================
# RSS 피드 소스
# ====================================================================
RSS_FEEDS = [
    {"name": "Reuters", "url": "https://feeds.reuters.com/reuters/businessNews", "lang": "en"},
    {"name": "CNBC", "url": "https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=100727362", "lang": "en"},
    {"name": "BBC", "url": "https://feeds.bbci.co.uk/news/business/rss.xml", "lang": "en"},
    {"name": "BBC World", "url": "https://feeds.bbci.co.uk/news/world/rss.xml", "lang": "en"},
    {"name": "연합뉴스(경제)", "url": "https://www.yna.co.kr/rss/economy.xml", "lang": "ko"},
    {"name": "연합뉴스(국제)", "url": "https://www.yna.co.kr/rss/international.xml", "lang": "ko"},
    # when:1d = 최근 1일 이내 뉴스만
    {"name": "Google News(경제)", "url": "https://news.google.com/rss/search?q=%EA%B8%B4%EA%B8%89+%EA%B2%BD%EC%A0%9C+%EC%86%8D%EB%B3%B4+when%3A1d&hl=ko&gl=KR&ceid=KR:ko", "lang": "ko"},
]

# ====================================================================
# 긴급 키워드 (카테고리별)
# ====================================================================
URGENT_KEYWORDS = {
    "시장_긴급": {
        "en": ["crash", "halt trading", "circuit breaker", "flash crash", "meltdown",
               "plunges", "plummets", "tumbles", "sell-off", "rout", "panic",
               "emergency", "record low", "worst day"],
        "ko": ["폭락", "서킷브레이커", "거래정지", "급락", "패닉", "공포", "대폭락",
               "사이드카", "긴급", "역대최저", "최악"],
    },
    "중앙은행_긴급": {
        "en": ["emergency rate cut", "emergency rate hike", "unscheduled meeting",
               "emergency fed", "surprise rate", "rate decision", "FOMC emergency",
               "BOJ intervention", "ECB emergency"],
        "ko": ["긴급 금리", "임시 FOMC", "긴급 인하", "긴급 인상", "중앙은행 개입",
               "긴급 회의", "금리 결정"],
    },
    "지정학_긴급": {
        "en": ["war breaks out", "strikes", "attacks", "invasion", "missile",
               "nuclear", "sanctions imposed", "ceasefire breaks",
               "airstrike", "bombing", "explosion"],
        "ko": ["전쟁 발발", "공습", "공격", "침공", "미사일", "핵", "제재",
               "휴전 결렬", "폭격", "폭발", "교전", "무력충돌"],
    },
    "기업_긴급": {
        "en": ["bankruptcy", "delisting", "Chapter 11", "trading suspended",
               "fraud", "CEO resigns", "insider trading"],
        "ko": ["파산", "상장폐지", "거래정지", "회계부정", "경영진 사임",
               "내부자거래", "분식회계"],
    },
    "원자재_긴급": {
        "en": ["oil surges", "gold record", "crude spike", "energy crisis"],
        "ko": ["유가 급등", "금 최고가", "에너지 위기", "원유 급등"],
    },
}

# 모든 키워드 플랫 리스트
ALL_KEYWORDS_EN = []
ALL_KEYWORDS_KO = []
for cat, data in URGENT_KEYWORDS.items():
    ALL_KEYWORDS_EN.extend(data["en"])
    ALL_KEYWORDS_KO.extend(data["ko"])


# ====================================================================
# .env 파싱
# ====================================================================
def parse_env_file(env_path):
    """config/.env 파일을 파싱합니다."""
    env_vars = {}
    if not os.path.exists(env_path):
        return env_vars
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if not value.startswith(("'", '"')) and "#" in value:
                value = value.split("#")[0].strip()
            if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
                value = value[1:-1]
            env_vars[key] = value
    return env_vars


# ====================================================================
# 텔레그램 전송
# ====================================================================
def telegram_send(bot_token, chat_id, text):
    """텔레그램 메시지 전송."""
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    data = urllib.parse.urlencode({
        "chat_id": chat_id,
        "text": text,
        "disable_web_page_preview": "false",
    }).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    with urllib.request.urlopen(req, timeout=10) as resp:
        result = json.loads(resp.read().decode("utf-8"))
        if not result.get("ok"):
            raise RuntimeError(f"Telegram API error: {result}")


# ====================================================================
# RSS 파싱
# ====================================================================
def fetch_rss(url, timeout=10):
    """RSS 피드를 가져와 뉴스 목록을 반환."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            content = resp.read()
    except Exception as e:
        print(f"  [RSS 실패] {url}: {e}")
        return []

    try:
        root = ET.fromstring(content)
    except ET.ParseError as e:
        print(f"  [XML 파싱 실패] {url}: {e}")
        return []

    items = []
    # RSS 2.0
    for item in root.iter("item"):
        title = _get_text(item, "title")
        link = _get_text(item, "link")
        pub = _get_text(item, "pubDate")
        if title and link:
            items.append({"title": title, "link": link, "pub": pub})
    # Atom
    for entry in root.iter("{http://www.w3.org/2005/Atom}entry"):
        title_el = entry.find("{http://www.w3.org/2005/Atom}title")
        link_el = entry.find("{http://www.w3.org/2005/Atom}link")
        pub_el = entry.find("{http://www.w3.org/2005/Atom}published")
        title = title_el.text if title_el is not None else ""
        link = link_el.get("href") if link_el is not None else ""
        pub = pub_el.text if pub_el is not None else ""
        if title and link:
            items.append({"title": title, "link": link, "pub": pub})

    return items


def _get_text(el, tag):
    """XML 요소에서 텍스트 추출."""
    child = el.find(tag)
    if child is None or child.text is None:
        return ""
    return child.text.strip()


def parse_pub_date(pub_str):
    """RSS pubDate 문자열을 datetime(UTC)으로 파싱.
    여러 포맷 지원: RFC822, ISO8601, etc. 실패 시 None."""
    if not pub_str:
        return None
    pub_str = pub_str.strip()

    # 시도할 포맷들
    formats = [
        "%a, %d %b %Y %H:%M:%S %z",       # RFC822: Mon, 19 Apr 2026 15:30:00 +0900
        "%a, %d %b %Y %H:%M:%S %Z",       # RFC822 with timezone name
        "%Y-%m-%dT%H:%M:%S%z",            # ISO8601: 2026-04-19T15:30:00+0900
        "%Y-%m-%dT%H:%M:%SZ",             # ISO8601 UTC: 2026-04-19T15:30:00Z
        "%Y-%m-%dT%H:%M:%S.%fZ",          # ISO8601 with ms
        "%Y-%m-%d %H:%M:%S",              # Simple format
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(pub_str, fmt)
            # timezone 없으면 UTC로 간주
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        except ValueError:
            continue

    # "GMT" 같은 문자열 치환 후 재시도
    cleaned = pub_str.replace(" GMT", " +0000").replace(" UT", " +0000")
    try:
        dt = datetime.strptime(cleaned, "%a, %d %b %Y %H:%M:%S %z")
        return dt.astimezone(timezone.utc)
    except ValueError:
        pass

    return None


def is_recent(pub_str, hours=6):
    """뉴스가 최근 N시간 이내에 발행되었는지 확인.
    pub_str 파싱 실패 시 True 반환 (관대한 처리)."""
    pub_dt = parse_pub_date(pub_str)
    if pub_dt is None:
        return True  # 파싱 실패 시 포함 (false positive 감수)
    now = datetime.now(timezone.utc)
    return (now - pub_dt) <= timedelta(hours=hours)


# ====================================================================
# 긴급 키워드 감지
# ====================================================================
def detect_urgent(title, lang="en"):
    """제목에서 긴급 키워드를 감지하고 카테고리 반환."""
    title_lower = title.lower()
    matched = []
    keywords_source = URGENT_KEYWORDS

    for category, data in keywords_source.items():
        kw_list = data["en"] if lang == "en" else data["ko"]
        # 반대 언어 키워드도 체크 (한국 기사에 영어 단어 포함 가능)
        other_list = data["ko"] if lang == "en" else data["en"]
        all_kw = kw_list + other_list
        for kw in all_kw:
            if kw.lower() in title_lower:
                matched.append((category, kw))
                break  # 카테고리당 하나만 매칭

    return matched


def category_emoji(category):
    """카테고리별 이모지."""
    return {
        "시장_긴급": "📉",
        "중앙은행_긴급": "🏦",
        "지정학_긴급": "⚠️",
        "기업_긴급": "🏢",
        "원자재_긴급": "🛢️",
    }.get(category, "📰")


def category_name(category):
    """카테고리 한글명."""
    return {
        "시장_긴급": "시장 긴급",
        "중앙은행_긴급": "중앙은행",
        "지정학_긴급": "지정학",
        "기업_긴급": "기업",
        "원자재_긴급": "원자재",
    }.get(category, "뉴스")


# ====================================================================
# 상태 관리 (이미 본 뉴스 추적)
# ====================================================================
def load_state():
    """이전 실행에서 본 뉴스 ID 목록을 로드."""
    if not os.path.exists(STATE_FILE):
        return {"seen_links": [], "last_updated": None}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"seen_links": [], "last_updated": None}


def save_state(seen_links):
    """현재 본 뉴스 링크를 저장 (최근 500개만 유지)."""
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    state = {
        "seen_links": list(seen_links)[-500:],
        "last_updated": datetime.now(timezone.utc).isoformat(),
    }
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 60)
    print("  긴급 뉴스 모니터링 시작")
    print(f"  시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # 텔레그램 설정 로드
    env_path = os.path.join(CONFIG_DIR, ".env")
    env_vars = parse_env_file(env_path)
    bot_token = env_vars.get("TELEGRAM_FINANCE_BOT_TOKEN") or os.environ.get("TELEGRAM_FINANCE_BOT_TOKEN")
    chat_id = env_vars.get("TELEGRAM_FINANCE_CHAT_ID") or os.environ.get("TELEGRAM_FINANCE_CHAT_ID")

    if not bot_token or not chat_id:
        print("  [경고] 텔레그램 토큰/chat_id 미설정. 콘솔 출력만 진행.")

    # 이전 상태 로드
    state = load_state()
    seen_links = set(state.get("seen_links", []))
    print(f"  이전 본 뉴스: {len(seen_links)}건")

    # RSS 피드 수집
    all_urgent = []
    for feed in RSS_FEEDS:
        print(f"\n  [{feed['name']}] 수집 중...")
        items = fetch_rss(feed["url"])
        print(f"    {len(items)}건 수신")

        for item in items:
            if item["link"] in seen_links:
                continue

            # 최근 6시간 이내 뉴스만 긴급 알림 대상
            # (pubDate 파싱 실패 시 관대하게 포함)
            if not is_recent(item.get("pub", ""), hours=6):
                seen_links.add(item["link"])  # 오래된 뉴스도 seen 처리
                continue

            matched = detect_urgent(item["title"], lang=feed["lang"])
            if matched:
                all_urgent.append({
                    "source": feed["name"],
                    "title": item["title"],
                    "link": item["link"],
                    "pub": item.get("pub", ""),
                    "matched": matched,
                    "lang": feed["lang"],
                })

            seen_links.add(item["link"])

    # 저장
    save_state(seen_links)

    print(f"\n{'=' * 60}")
    print(f"  긴급 뉴스 감지: {len(all_urgent)}건")
    print("=" * 60)

    if not all_urgent:
        print("  긴급 뉴스 없음 (알림 전송 안 함)")
        return 0

    # 텔레그램 전송 (카테고리별 그룹화)
    if bot_token and chat_id:
        # 카테고리별로 정리
        by_category = {}
        for news in all_urgent:
            for cat, kw in news["matched"]:
                if cat not in by_category:
                    by_category[cat] = []
                by_category[cat].append(news)
                break  # 한 뉴스는 한 카테고리만

        # 메시지 조립
        priority_order = ["시장_긴급", "중앙은행_긴급", "지정학_긴급", "원자재_긴급", "기업_긴급"]
        lines = ["🚨 긴급 뉴스 알림", "=" * 25, ""]

        max_items = 15  # 텔레그램 메시지 길이 제한
        total_shown = 0

        for cat in priority_order:
            if cat not in by_category:
                continue
            emoji = category_emoji(cat)
            lines.append(f"{emoji} {category_name(cat)}")
            for news in by_category[cat][:5]:
                if total_shown >= max_items:
                    break
                title = news["title"][:100]
                lines.append(f"  • [{news['source']}] {title}")
                lines.append(f"    {news['link']}")
                total_shown += 1
            lines.append("")

        if total_shown < len(all_urgent):
            lines.append(f"... 외 {len(all_urgent) - total_shown}건")

        lines.append(f"\n검사 시각: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

        message = "\n".join(lines)

        # 텔레그램 제한: 4096자
        if len(message) > 4000:
            message = message[:4000] + "\n\n... (추가 뉴스는 대시보드 참조)"

        try:
            telegram_send(bot_token, chat_id, message)
            print(f"  텔레그램 전송 완료: {total_shown}건")
        except Exception as e:
            print(f"  [텔레그램 실패] {e}")
            return 1
    else:
        # 콘솔에만 출력
        for news in all_urgent:
            cats = ", ".join(c for c, _ in news["matched"])
            print(f"  [{cats}] [{news['source']}] {news['title']}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
