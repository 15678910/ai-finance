"""
CLARITY Act 및 암호화폐 입법 패키지 실시간 모니터
===================================================

추적 대상 법안 (미국 의회):
  - HR 3633: CLARITY Act (Digital Asset Market Clarity Act) — 119대
  - S 1582:  GENIUS Act (Stablecoin Innovation Act) — 119대 (2025.7 서명 완료)
  - HR 4763: FIT21 (Financial Innovation Technology for 21C) — 118대
  - (Lummis-Gillibrand 등은 별도 ID 확인되면 추가)

알림 기준 (옵션 B — 주요 이벤트만):
  - 위원회 표결 통과 (Reported)
  - 본회의 표결 통과 (Passed House / Passed Senate)
  - 대통령 서명 / 거부 (Enacted / Vetoed)

추가 RSS 피드 (법안 관련 키워드 필터):
  - CoinDesk, The Block, Decrypt

데이터 소스:
  - GovTrack API (무료, 키 불필요): https://www.govtrack.us/developers/api
  - Congress.gov API (선택, 더 정밀): CONGRESS_API_KEY 환경변수 설정 시

빈도: 4시간 (GH Actions cron)

🚨 정보 모니터링 도구. 투자 결정에 단독 사용 금지.
"""

import os
import json
import re
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta

try:
    from defusedxml import ElementTree as ET  # type: ignore
except ImportError:
    import xml.etree.ElementTree as ET  # XXE 위험 — defusedxml 권장

from core import get_secret, send_message, load_state, save_state

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "clarity_act.json")
STATE_NAME = "clarity_act"
KST = timezone(timedelta(hours=9))

USER_AGENT = "ai-finance-clarity-monitor/1.0 (https://github.com/15678910/ai-finance)"
TIMEOUT_SEC = 15

# ====================================================================
# 추적 대상 법안 (GovTrack 식별자)
# ====================================================================
TRACKED_BILLS = [
    {
        "short_name": "CLARITY Act",
        "full_name": "Digital Asset Market Clarity Act",
        "congress": 119, "bill_type": "house_bill", "number": 3633,
        "category": "core",
    },
    {
        "short_name": "GENIUS Act",
        "full_name": "Guiding and Establishing National Innovation for US Stablecoins Act",
        "congress": 119, "bill_type": "senate_bill", "number": 1582,
        "category": "core",
    },
    {
        "short_name": "FIT21",
        "full_name": "Financial Innovation and Technology for the 21st Century Act",
        "congress": 118, "bill_type": "house_bill", "number": 4763,
        "category": "related",
    },
    # 추가 가능: Lummis-Gillibrand 재발의 시 등록
]

# 상태 코드 매핑 — "주요 이벤트" 여부 (emoji, label_kr 순)
# https://www.govtrack.us/developers/api → BillStatus
MAJOR_STATUS_CODES = {
    "reported": ("🟡", "위원회 표결 통과"),
    "pass_over_house": ("🟢", "하원 본회의 통과"),
    "pass_over_senate": ("🟢", "상원 본회의 통과"),
    "passed_bill": ("🟢", "양원 모두 통과"),
    "pass_back_house": ("🟢", "하원 재의결"),
    "pass_back_senate": ("🟢", "상원 재의결"),
    "conference_passed_house": ("🟢", "협의위원회안 하원 통과"),
    "conference_passed_senate": ("🟢", "협의위원회안 상원 통과"),
    "enacted_signed": ("✅", "대통령 서명 → 법률 제정"),
    "enacted_veto_override": ("✅", "거부권 무효화 (양원 2/3)"),
    "prov_kill_veto": ("🔴", "대통령 거부권 행사"),
    "vetoed_pocket": ("🔴", "Pocket Veto (회기 종료 후 묵살)"),
    "vetoed_override_fail_house": ("🔴", "거부권 무효화 실패 (하원)"),
    "vetoed_override_fail_senate": ("🔴", "거부권 무효화 실패 (상원)"),
    "fail_originating_house": ("🔴", "하원 부결"),
    "fail_originating_senate": ("🔴", "상원 부결"),
}

# ====================================================================
# 암호화폐 입법 뉴스 RSS
# ====================================================================
CRYPTO_LEG_FEEDS = [
    {"name": "CoinDesk", "url": "https://www.coindesk.com/arc/outboundfeeds/rss/"},
    {"name": "The Block", "url": "https://www.theblock.co/rss.xml"},
    {"name": "Decrypt", "url": "https://decrypt.co/feed"},
]

LEGISLATION_KEYWORDS = [
    # 법안 이름
    "CLARITY Act", "Digital Asset Market Clarity",
    "GENIUS Act", "Stablecoin Innovation",
    "FIT21", "Financial Innovation and Technology",
    "Lummis-Gillibrand", "Responsible Financial Innovation",
    # 입법 일반
    "stablecoin bill", "stablecoin act", "stablecoin regulation",
    "crypto bill", "crypto legislation", "crypto regulation",
    "Senate Banking Committee", "House Financial Services",
    "SEC vs", "CFTC oversight",
    "executive order crypto", "crypto framework",
    # 표결·서명 관련
    "passes Senate crypto", "passes House crypto",
    "signed into law crypto", "vetoes crypto",
]


# ====================================================================
# HTTP 헬퍼
# ====================================================================
def http_get(url: str, headers: dict | None = None, timeout: int = TIMEOUT_SEC) -> bytes:
    """HTTP GET → bytes. 표준 라이브러리만 사용."""
    req_headers = {"User-Agent": USER_AGENT}
    if headers:
        req_headers.update(headers)
    req = urllib.request.Request(url, headers=req_headers)
    with urllib.request.urlopen(req, timeout=timeout) as resp:  # nosec — RSS/공개 API
        return resp.read()


# ====================================================================
# 법안 상태 조회 — Congress.gov API (1차) / GovTrack (2차 폴백)
# ====================================================================
# Congress.gov "latestAction"에서 우리가 "주요 이벤트"로 간주할 패턴
CONGRESS_ACTION_PATTERNS = [
    (r"^Became Public Law", "enacted_signed"),
    (r"signed by president", "enacted_signed"),
    (r"presented to president", "passed_bill"),
    (r"passed/agreed to in (?:the )?house", "pass_over_house"),
    (r"passed (?:the )?house", "pass_over_house"),
    (r"on passage passed.*house", "pass_over_house"),
    (r"passed/agreed to in (?:the )?senate", "pass_over_senate"),
    (r"passed (?:the )?senate", "pass_over_senate"),
    (r"on passage.*senate", "pass_over_senate"),
    (r"reported by", "reported"),
    (r"ordered to be reported", "reported"),
    (r"committee.*vote.*favorable", "reported"),
    (r"vetoed by president", "prov_kill_veto"),
    (r"failed.*passage", "fail_originating_house"),
]


def classify_action(text: str) -> tuple[str | None, str]:
    """Congress.gov latestAction 텍스트 → (status_code, label_kr) 매핑.

    re.IGNORECASE 사용 — 패턴 대소문자와 무관하게 매칭.
    """
    if not text:
        return None, ""
    for pattern, code in CONGRESS_ACTION_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            # MAJOR_STATUS_CODES 값: (emoji, label_kr) — 한국어 라벨은 인덱스 [1]
            label_kr = MAJOR_STATUS_CODES.get(code, ("", ""))[1]
            return code, label_kr
    return None, ""


def fetch_bill_congress_gov(bill: dict, api_key: str) -> dict | None:
    """Congress.gov API에서 법안 정보 조회 (공식 1차 데이터)."""
    bt_map = {"house_bill": "hr", "senate_bill": "s",
              "house_joint_resolution": "hjres", "senate_joint_resolution": "sjres"}
    bt = bt_map.get(bill["bill_type"])
    if not bt:
        return None
    url = (f"https://api.congress.gov/v3/bill/{bill['congress']}/{bt}/{bill['number']}"
           f"?format=json&api_key={api_key}")
    try:
        raw = http_get(url, headers={"Accept": "application/json"})
        data = json.loads(raw.decode("utf-8"))
    except Exception as e:
        print(f"    [Congress.gov 실패] {e}")
        return None

    b = data.get("bill", {})
    latest = b.get("latestAction", {}) or {}
    action_text = latest.get("text", "")
    action_date = latest.get("actionDate", "")

    code, _ = classify_action(action_text)  # label_kr은 호출부에서 재조회
    sponsors = b.get("sponsors", []) or []
    sponsor_name = sponsors[0].get("fullName") if sponsors else None

    link = f"https://www.congress.gov/bill/{bill['congress']}th-congress/" + \
           ("house-bill/" if bt == "hr" else "senate-bill/") + str(bill["number"])

    return {
        "id": f"{bill['congress']}-{bt}-{bill['number']}",
        "title": b.get("title", ""),
        "status_code": code,
        "status_label": action_text,
        "status_date": action_date,
        "introduced_date": b.get("introducedDate"),
        "link": link,
        "is_alive": code not in ("enacted_signed", "prov_kill_veto", "fail_originating_house",
                                 "fail_originating_senate", "vetoed_pocket"),
        "sponsor": sponsor_name,
        "source": "Congress.gov",
    }


def fetch_bill_govtrack(bill: dict) -> dict | None:
    """GovTrack 폴백 (User-Agent 명시)."""
    params = urllib.parse.urlencode({
        "congress": bill["congress"],
        "bill_type": bill["bill_type"],
        "number": bill["number"],
    })
    url = f"https://www.govtrack.us/api/v2/bill?{params}"
    try:
        raw = http_get(url, headers={"Accept": "application/json"})
        data = json.loads(raw.decode("utf-8"))
    except Exception as e:
        print(f"    [GovTrack 실패] {e}")
        return None

    objs = data.get("objects", [])
    if not objs:
        return None
    obj = objs[0]
    return {
        "id": obj.get("id"),
        "title": obj.get("title", "") or obj.get("title_without_number", ""),
        "status_code": obj.get("current_status"),
        "status_label": obj.get("current_status_description"),
        "status_date": obj.get("current_status_date"),
        "introduced_date": obj.get("introduced_date"),
        "link": obj.get("link"),
        "is_alive": obj.get("is_alive"),
        "sponsor": (obj.get("sponsor") or {}).get("name"),
        "source": "GovTrack",
    }


def fetch_bill_status(bill: dict, api_key_warned: list) -> dict | None:
    """Congress.gov 1차 → GovTrack 폴백."""
    api_key = get_secret("CONGRESS_API_KEY")
    if api_key:
        result = fetch_bill_congress_gov(bill, api_key)
        if result:
            return result
        print("    Congress.gov 실패 → GovTrack 폴백 시도")
    elif not api_key_warned:
        print("    ⚠️  CONGRESS_API_KEY 미설정 → GovTrack 폴백 시도")
        print("       (Congress.gov 무료 키: https://api.congress.gov/sign-up/)")
        api_key_warned.append(True)
    return fetch_bill_govtrack(bill)


# ====================================================================
# RSS 파싱 (법안 키워드 필터)
# ====================================================================
def matches_legislation(text: str) -> str | None:
    """법안 관련 키워드 매칭. 매칭된 키워드 반환 또는 None."""
    if not text:
        return None
    text_low = text.lower()
    for kw in LEGISLATION_KEYWORDS:
        if kw.lower() in text_low:
            return kw
    return None


def parse_rss_items(xml_bytes: bytes) -> list:
    """RSS XML → [{title, link, summary, pub_date, guid}]"""
    items = []
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        print(f"  [RSS 파싱 실패] {e}")
        return items

    # RSS 2.0 (channel/item)
    for item in root.iter("item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        summary = (item.findtext("description") or "").strip()
        # 태그 제거
        summary = re.sub(r"<[^>]+>", "", summary)
        pub_date = (item.findtext("pubDate") or "").strip()
        guid = (item.findtext("guid") or link).strip()
        if title:
            items.append({
                "title": title, "link": link, "summary": summary[:300],
                "pub_date": pub_date, "guid": guid,
            })

    # Atom (feed/entry) — Decrypt, 일부 피드
    ns = "{http://www.w3.org/2005/Atom}"
    for entry in root.iter(f"{ns}entry"):
        title_el = entry.find(f"{ns}title")
        link_el = entry.find(f"{ns}link")
        summary_el = entry.find(f"{ns}summary") or entry.find(f"{ns}content")
        pub_el = entry.find(f"{ns}published") or entry.find(f"{ns}updated")
        id_el = entry.find(f"{ns}id")

        title = ((title_el.text or "") if title_el is not None else "").strip()
        link = ""
        if link_el is not None:
            link = link_el.get("href") or (link_el.text or "")
        summary = ((summary_el.text or "") if summary_el is not None else "").strip()
        summary = re.sub(r"<[^>]+>", "", summary)
        pub_date = ((pub_el.text or "") if pub_el is not None else "").strip()
        guid = ((id_el.text or link) if id_el is not None else link).strip()

        if title:
            items.append({
                "title": title, "link": link, "summary": summary[:300],
                "pub_date": pub_date, "guid": guid,
            })

    return items


def fetch_rss_legislation_news(state: dict) -> list:
    """RSS에서 법안 관련 신규 뉴스 추출. 이미 알린 guid는 제외."""
    alerted_guids = set(state.get("alerted_news_guids", []))
    new_news = []

    for feed in CRYPTO_LEG_FEEDS:
        print(f"  RSS: {feed['name']}...")
        try:
            raw = http_get(feed["url"])
        except Exception as e:
            print(f"    [실패] {e}")
            continue

        items = parse_rss_items(raw)
        matched = 0
        for it in items:
            haystack = f"{it['title']} {it['summary']}"
            kw = matches_legislation(haystack)
            if not kw:
                continue
            if it["guid"] in alerted_guids:
                continue
            it["matched_keyword"] = kw
            it["source"] = feed["name"]
            new_news.append(it)
            alerted_guids.add(it["guid"])
            matched += 1
        print(f"    {len(items)}건 중 법안 관련 신규 {matched}건")

    # state 갱신 (alerted_guids 트림 — 최대 500개)
    state["alerted_news_guids"] = list(alerted_guids)[-500:]
    return new_news


# ====================================================================
# 텔레그램 알림 포맷
# ====================================================================
def format_bill_change_alert(bill: dict, prev: dict | None, curr: dict) -> str:
    """법안 상태 변경 알림 메시지."""
    emoji, label_kr = MAJOR_STATUS_CODES.get(curr["status_code"], ("📋", curr.get("status_label", "상태 변경")))
    lines = [
        f"{emoji} 미국 입법 추적 — 주요 이벤트",
        "=" * 30,
        f"법안: {bill['short_name']}",
        f"     ({bill['full_name']})",
        f"의회: {bill['congress']}대 · {bill['bill_type'].replace('_', ' ').title()} {bill['number']}",
        "",
        f"새 상태: {label_kr}",
        f"공식 라벨: {curr.get('status_label', '—')}",
        f"발생일: {curr.get('status_date', '—')}",
    ]
    if prev and prev.get("status_code"):
        prev_label = MAJOR_STATUS_CODES.get(prev["status_code"], ("📋", prev.get("status_label", "—")))[1]
        lines.append(f"이전 상태: {prev_label}")
    if curr.get("sponsor"):
        lines.append(f"발의자: {curr['sponsor']}")
    if curr.get("link"):
        lines.append(f"\n🔗 GovTrack: {curr['link']}")
    lines.append("\n🚨 정보 모니터링 도구. 투자 결정 단독 사용 금지.")
    lines.append("대시보드: https://15678910.github.io/ai-finance/")
    return "\n".join(lines)


def format_news_digest(news: list) -> str:
    """RSS 법안 뉴스 다이제스트."""
    lines = ["📰 암호화폐 입법 뉴스 (법안 관련 키워드 매칭)", "=" * 30, ""]
    for n in news[:8]:  # 한 메시지에 최대 8건
        lines.append(f"🔸 [{n['source']}] {n['title']}")
        if n.get("matched_keyword"):
            lines.append(f"   키워드: {n['matched_keyword']}")
        if n.get("link"):
            lines.append(f"   {n['link']}")
        lines.append("")
    if len(news) > 8:
        lines.append(f"... 외 {len(news) - 8}건")
    lines.append("🚨 정보 모니터링. 자동 매매 금지.")
    lines.append("대시보드: https://15678910.github.io/ai-finance/")
    return "\n".join(lines)


# ====================================================================
# 메인
# ====================================================================
def main():
    print("=" * 70)
    print("  CLARITY Act 및 암호화폐 입법 모니터")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    state = load_state(STATE_NAME, default={"bills": {}, "alerted_news_guids": []})

    # 1) 법안 상태 추적
    print("\n[법안 상태 조회 — GovTrack API]")
    bill_results = []
    alerts_to_send = []
    api_key_warned: list = []

    for bill in TRACKED_BILLS:
        bill_key = f"{bill['congress']}-{bill['bill_type']}-{bill['number']}"
        print(f"  {bill['short_name']} (H/S {bill['number']})...")
        curr = fetch_bill_status(bill, api_key_warned)
        if not curr:
            continue

        prev = state.get("bills", {}).get(bill_key)
        status_code = curr.get("status_code")
        is_major = status_code in MAJOR_STATUS_CODES
        is_new_status = (not prev) or (prev.get("status_code") != status_code) or \
                        (prev.get("status_date") != curr.get("status_date"))

        if status_code and status_code in MAJOR_STATUS_CODES:
            emoji, label_kr = MAJOR_STATUS_CODES[status_code]
        else:
            emoji, label_kr = ("📋", curr.get("status_label") or "—")
        print(f"    상태: {emoji} {label_kr}  ({curr.get('status_date', '—')})")

        # 주요 이벤트 + 새 변경 → 알림 대상
        if is_major and is_new_status:
            msg = format_bill_change_alert(bill, prev, curr)
            alerts_to_send.append(msg)
            print(f"    🚨 알림 대상 (이전 상태 대비 변화)")

        # 상태 저장
        state.setdefault("bills", {})[bill_key] = curr

        bill_results.append({
            "short_name": bill["short_name"],
            "full_name": bill["full_name"],
            "congress": bill["congress"],
            "bill_type": bill["bill_type"],
            "number": bill["number"],
            "category": bill["category"],
            "status_code": status_code,
            "status_label": curr.get("status_label"),
            "status_label_kr": label_kr,
            "status_date": curr.get("status_date"),
            "status_emoji": emoji,
            "is_major_event": is_major,
            "sponsor": curr.get("sponsor"),
            "link": curr.get("link"),
            "is_alive": curr.get("is_alive"),
        })

    # 2) RSS 뉴스 추적
    print("\n[암호화폐 입법 뉴스 RSS — 키워드 필터]")
    news = fetch_rss_legislation_news(state)
    if news:
        alerts_to_send.append(format_news_digest(news))

    # 3) 알림 발송
    print(f"\n[알림 발송] {len(alerts_to_send)}건")
    for msg in alerts_to_send:
        send_message(msg)
        print("  ✅ 텔레그램 발송")

    if not alerts_to_send:
        print("  변경 사항 없음 — 알림 미발송")

    # 4) 상태 + 결과 저장
    save_state(STATE_NAME, state)

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "tracked_bills_count": len(bill_results),
        "alerts_sent_this_run": len(alerts_to_send),
        "bills": bill_results,
        "recent_news": news[:20],  # 최근 20건 대시보드용
        "status_legend": {code: label for code, (_, label) in MAJOR_STATUS_CODES.items()},
        "data_sources": [
            "GovTrack API (www.govtrack.us/developers/api)",
            "CoinDesk RSS", "The Block RSS", "Decrypt RSS",
        ],
        "frequency": "4시간 (GH Actions cron)",
        "warning": "🚨 정보 모니터링 도구. 투자 결정에 단독 사용 금지.",
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n  결과 저장: {OUTPUT_FILE}")

    print("\n" + "=" * 70)
    print(f"  완료: 법안 {len(bill_results)}개 · 뉴스 신규 {len(news)}건 · 알림 {len(alerts_to_send)}건")
    print("=" * 70)


if __name__ == "__main__":
    main()
