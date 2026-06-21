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
try:
    from defusedxml import ElementTree as ET  # type: ignore
except ImportError:
    import xml.etree.ElementTree as ET  # XXE/billion-laughs 위험. defusedxml 권장.
from datetime import datetime, timezone, timedelta
from pathlib import Path

from core import get_secret, send_message, load_state, save_state
from news_impact import classify_news

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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
    # 이공계 석박사 전문 기술 분석 (반도체, AI, 바이오, 핵융합 등)
    {"name": "Heisenberg(기술분석)", "url": "https://heisenberg.kr/feed/", "lang": "ko"},
    # BOJ 사전신호 — 닛케이/로이터/블룸버그 관측보도 선점(회의 며칠 전 소식통 인용으로 흘러나옴)
    {"name": "Google News(BOJ)", "url": "https://news.google.com/rss/search?q=%22Bank+of+Japan%22+OR+BOJ+OR+Ueda+rate+when%3A2d&hl=en-US&gl=US&ceid=US:en", "lang": "en"},
    {"name": "Google News(일본은행)", "url": "https://news.google.com/rss/search?q=%EC%9D%BC%EB%B3%B8%EC%9D%80%ED%96%89+%EA%B8%88%EB%A6%AC+when%3A2d&hl=ko&gl=KR&ceid=KR:ko", "lang": "ko"},
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
    # BOJ 사전신호(관측보도) — 최우선·쿨다운 없음. 회의 전 흘러나오는 방향 기사 선점.
    "중앙은행_관측": {
        "en": ["boj to raise", "boj to hike", "boj to consider", "boj said to", "boj mulls",
               "boj to hold", "boj to keep", "boj eyes", "boj weighs", "boj to debate",
               "boj likely", "boj signals", "boj to lift", "boj hikes", "boj raises",
               "boj holds", "boj keeps", "boj lifts", "boj meeting", "boj policy",
               "boj rate hike", "boj rate decision", "ueda",
               "boj set to", "boj set for", "boj poised", "boj to lift rate",
               "bank of japan to", "bank of japan set to", "bank of japan poised",
               "bank of japan likely", "bank of japan rate",
               "bank of japan hikes", "bank of japan raises", "bank of japan holds",
               # 회견 2차 신호(포워드 가이던스·QT) — BOJ/우에다 문맥 한정
               "boj more hikes", "boj further hike", "boj additional hike",
               "boj neutral rate", "boj rate path", "boj to continue", "boj signals more",
               "boj qt", "boj taper", "boj to taper", "boj jgb", "boj bond purchase",
               "ueda hawkish", "ueda dovish", "ueda neutral", "ueda signals"],
        "ko": ["일본은행 인상", "일본은행 금리", "일본은행 동결", "일본은행 회의", "일본은행 검토",
               "boj 인상", "boj 금리", "우에다", "日銀", "利上げ", "据え置き",
               "金融政策決定会合", "엔캐리 청산", "엔 캐리 청산",
               # 회견 2차 신호(일본어·일본 문맥은 자체로 BOJ 한정이라 안전)
               "追加利上げ", "中立金利", "連続利上げ", "量的引き締め", "国債買い入れ",
               "일본은행 추가", "일본 중립금리", "우에다 매파", "우에다 비둘기"],
    },
    # 메모리·SK실적 변곡 — 주가가 메모리가격·실적을 1~2분기 선행하므로 가격전환·감산·실적을 최우선 감시.
    "메모리_실적": {
        "en": ["dram price", "nand price", "hbm price", "memory price", "memory prices",
               "dram contract", "dram spot", "memory chip price", "chip price hike",
               "memory glut", "memory shortage", "production cut", "output cut", "capacity cut",
               "sk hynix earnings", "sk hynix profit", "sk hynix results", "sk hynix record",
               "sk hynix", "hynix earnings", "hbm earnings", "record profit", "record earnings",
               "record operating profit", "posts record", "earnings beat", "profit surge",
               "micron earnings", "micron results", "micron guidance", "samsung memory",
               "hbm demand", "hbm sold out", "hbm capacity", "dram inventory",
               "memory downturn", "memory upcycle", "memory supercycle", "dxi index"],
        "ko": ["디램 가격", "디램값", "낸드 가격", "낸드값", "메모리 가격", "메모리값",
               "메모리 가격 인상", "메모리 가격 하락", "고정거래가", "현물가격", "감산", "증산",
               "재고 소진", "재고 급증", "sk하이닉스 실적", "sk하이닉스 영업이익", "하이닉스 어닝",
               "하이닉스 최대 실적", "마이크론 실적", "삼성전자 메모리", "hbm 수요", "hbm 완판",
               "hbm 증설", "메모리 다운사이클", "메모리 업사이클", "디램 사이클", "메모리 슈퍼사이클"],
    },
    "중앙은행_긴급": {
        "en": ["emergency rate cut", "emergency rate hike", "unscheduled meeting",
               "emergency fed", "surprise rate", "rate decision", "FOMC emergency",
               "BOJ intervention", "ECB emergency"],
        "ko": ["긴급 금리", "임시 FOMC", "긴급 인하", "긴급 인상", "중앙은행 개입",
               "긴급 회의", "금리 결정"],
    },
    "환율_긴급": {
        "en": ["yen plunges", "yen tumbles", "yen intervention", "FX intervention",
               "currency intervention", "yen weakens sharply", "BOJ intervenes",
               "won plunges", "currency crisis", "dollar-yen surges", "yen slides"],
        "ko": ["엔화 급락", "엔화 약세", "엔화 급등", "달러엔 급등", "엔저", "슈퍼 엔저",
               "외환시장 개입", "환율 급등", "환율 급변", "환율 개입", "일본은행 개입",
               "원화 급락", "원화 약세", "외환 위기", "엔 매수 개입", "스무딩 오퍼레이션"],
    },
    "신용_긴급": {
        "en": ["redemption halt", "redemption suspended", "redemptions frozen", "fund gating",
               "margin call", "margin calls", "fire sale", "fire-sale", "forced selling",
               "distressed selling", "credit line cut", "credit crunch", "liquidity crisis",
               "bank run", "default", "debt default", "capital call default", "LP default",
               "fund freezes", "halts withdrawals", "credit event", "deleveraging"],
        "ko": ["환매 중단", "환매 연기", "환매 정지", "마진콜", "마진 콜", "헐값 매각",
               "투매", "강제 매각", "신용한도 축소", "여신 축소", "만기연장 거부",
               "디폴트", "채무불이행", "유동성 위기", "유동성 경색", "신용경색", "자금경색",
               "뱅크런", "펀드런", "환매 폭주", "부실 매각", "신용 이벤트", "디레버리징"],
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
    "비트코인_정책": {
        "en": ["Bitcoin Reserve", "Strategic Bitcoin", "Bitcoin Standard",
               "ARMA Act", "GENIUS Act", "Bitcoin Act", "Lummis bill",
               "Stablecoin regulation", "Compute dollar", "SDAC",
               "Bitcoin treasury", "Petro-dollar end", "digital dollar",
               "national Bitcoin", "kill switch stablecoin"],
        "ko": ["비트코인 비축", "비트코인 본위제", "준비금 현대화법",
               "스테이블코인 규제", "컴퓨트 달러", "전략 비트코인",
               "디지털 달러", "비트코인 법안", "암호화폐 패권",
               "비트코인 국가", "스테이블코인 의무"],
    },
    "기술_트렌드": {
        "en": ["token shortage", "HBM shortage", "EUV", "foundry capacity",
               "CXL memory", "quantum supremacy", "fusion ignition",
               "humanoid robot", "virtual cell", "foundation model"],
        "ko": ["토큰 쇼티지", "HBM 부족", "파운드리 증설", "EUV",
               "양자컴퓨터 상용화", "핵융합 점화", "토카막",
               "휴머노이드", "가상세포", "파운데이션 모델",
               "AI 인프라", "GPU 부족"],
    },
}

# 모든 키워드 플랫 리스트
ALL_KEYWORDS_EN = []
ALL_KEYWORDS_KO = []
for cat, data in URGENT_KEYWORDS.items():
    ALL_KEYWORDS_EN.extend(data["en"])
    ALL_KEYWORDS_KO.extend(data["ko"])


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
        "중앙은행_관측": "🔭",
        "메모리_실적": "💾",
        "시장_긴급": "📉",
        "중앙은행_긴급": "🏦",
        "환율_긴급": "💱",
        "신용_긴급": "💳",
        "지정학_긴급": "⚠️",
        "기업_긴급": "🏢",
        "원자재_긴급": "🛢️",
    }.get(category, "📰")


def category_name(category):
    """카테고리 한글명."""
    return {
        "중앙은행_관측": "🇯🇵 BOJ 관측·사전신호",
        "메모리_실적": "💾 메모리·SK 실적 변곡",
        "시장_긴급": "시장 긴급",
        "중앙은행_긴급": "중앙은행",
        "환율_긴급": "환율/외환",
        "신용_긴급": "신용/유동성",
        "지정학_긴급": "지정학",
        "기업_긴급": "기업",
        "원자재_긴급": "원자재",
    }.get(category, "뉴스")


# 카테고리별 알림 쿨다운(시간) — 큰 이벤트 직전 같은 주제 긴급뉴스 도배 방지.
# 시장_긴급·신용_긴급은 최우선이라 쿨다운 없음(항상 즉시 발송).
COOLDOWN_HOURS = {
    "메모리_실적": 2,
    "환율_긴급": 4,
    "중앙은행_긴급": 3,
    "지정학_긴급": 3,
    "원자재_긴급": 4,
    "기업_긴급": 3,
}


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 60)
    print("  긴급 뉴스 모니터링 시작")
    print(f"  시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # 텔레그램 설정 로드
    bot_token = get_secret("TELEGRAM_FINANCE_BOT_TOKEN")
    chat_id = get_secret("TELEGRAM_FINANCE_CHAT_ID")

    if not bot_token or not chat_id:
        print("  [경고] 텔레그램 토큰/chat_id 미설정. 콘솔 출력만 진행.")

    # 이전 상태 로드
    state = load_state("breaking_news", default={"seen_links": [], "last_updated": None})
    seen_links = set(state.get("seen_links", []))
    cat_last_sent = dict(state.get("cat_last_sent", {}))  # {카테고리: 마지막 발송 ISO ts}
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
    save_state("breaking_news", {
        "seen_links": list(seen_links)[-500:],
        "last_updated": datetime.now(timezone.utc).isoformat(),
    })

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

        # 카테고리별 쿨다운 적용 — 직전 발송 후 쿨다운 시간 내면 이번 회차에서 제외
        now_utc = datetime.now(timezone.utc)

        def _on_cooldown(cat):
            h = COOLDOWN_HOURS.get(cat)
            if not h:
                return False  # 시장·신용 등 최우선은 항상 발송
            ts = cat_last_sent.get(cat)
            if not ts:
                return False
            try:
                last = datetime.fromisoformat(ts)
            except Exception:
                return False
            return (now_utc - last) < timedelta(hours=h)

        suppressed = [c for c in list(by_category) if _on_cooldown(c)]
        for c in suppressed:
            del by_category[c]
        if suppressed:
            print(f"  쿨다운 보류 카테고리: {', '.join(suppressed)}")
        if not by_category:
            print("  모든 긴급 카테고리가 쿨다운 중 → 알림 전송 안 함")
            save_state("breaking_news", {
                "seen_links": list(seen_links)[-500:],
                "cat_last_sent": cat_last_sent,
                "last_updated": now_utc.isoformat(),
            })
            return 0

        # 메시지 조립
        priority_order = ["중앙은행_관측", "메모리_실적", "시장_긴급", "신용_긴급", "중앙은행_긴급", "환율_긴급", "지정학_긴급", "원자재_긴급", "기업_긴급"]
        lines = ["🚨 긴급 뉴스 알림", "=" * 25, ""]

        max_items = 15  # 텔레그램 메시지 길이 제한
        total_shown = 0
        shown_cats = set()  # 이번 회차에 실제로 발송된 카테고리(쿨다운 갱신용)

        for cat in priority_order:
            if cat not in by_category:
                continue
            emoji = category_emoji(cat)
            lines.append(f"{emoji} {category_name(cat)}")
            cnt_before = total_shown
            for news in by_category[cat][:5]:
                if total_shown >= max_items:
                    break
                title = news["title"][:100]
                _, _, semoji, impact = classify_news(news["title"])
                tag = f" {semoji}{impact}" if impact else ""
                lines.append(f"  • [{news['source']}] {title}{tag}")
                lines.append(f"    {news['link']}")
                total_shown += 1
            if total_shown > cnt_before:
                shown_cats.add(cat)
            lines.append("")

        if total_shown < len(all_urgent):
            lines.append(f"... 외 {len(all_urgent) - total_shown}건")

        lines.append(f"\n검사 시각: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

        message = "\n".join(lines)

        # 텔레그램 제한: 4096자
        if len(message) > 4000:
            message = message[:4000] + "\n\n... (추가 뉴스는 대시보드 참조)"

        ok = send_message(message, disable_preview=False)
        if ok:
            print(f"  텔레그램 전송 완료: {total_shown}건")
            # 실제 발송된 카테고리만 쿨다운 타이머 갱신 후 저장
            for c in shown_cats:
                if c in COOLDOWN_HOURS:
                    cat_last_sent[c] = now_utc.isoformat()
            save_state("breaking_news", {
                "seen_links": list(seen_links)[-500:],
                "cat_last_sent": cat_last_sent,
                "last_updated": now_utc.isoformat(),
            })
        else:
            print("  [텔레그램 실패]")
            return 1
    else:
        # 콘솔에만 출력
        for news in all_urgent:
            cats = ", ".join(c for c, _ in news["matched"])
            print(f"  [{cats}] [{news['source']}] {news['title']}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
