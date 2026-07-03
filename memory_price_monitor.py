"""
메모리 가격 뉴스 자동 추적 — DRAM·HBM·DDR5·LPDDR·NAND
========================================================
DRAM 스팟가·DXI는 TrendForce 유료라 무단 수집 불가. 대신 '메모리 가격 뉴스'를
무료(Google News RSS)로 자동 수집·방향 태깅해 가격 모멘텀을 실시간 근사 추적.
가격 인상/하락 뉴스 신규 발생 시 텔레그램 경보. HBM은 실시간 가격 없음(계약 기반)이라 뉴스로만.

출력: docs/memory_price.json
🚨 뉴스 기반 근사. 실제 계약가·스팟가는 TrendForce 등 원문 확인. 투자자문 아님.
"""

import json
import os
import re
import sys
import html as _html
import urllib.request
import urllib.parse
from email.utils import parsedate_to_datetime
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "memory_price.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

# 검색 질의(한국어 중심 + 영문 보강)
QUERIES = [
    "D램 가격 OR DRAM 가격 OR 메모리 반도체 가격",
    "HBM 가격 OR HBM 수요 OR HBM 증설 OR HBM ASP",
    "DDR5 가격 OR LPDDR 가격 OR 모바일 D램",
    "낸드 가격 OR NAND price OR 메모리 고정거래가 OR D램 스팟",
    "DRAM price OR memory chip price TrendForce",
]

# 제목 필터: 메모리 키워드 AND 가격/수급 키워드 (노이즈 제거)
MEM_KW = re.compile(r"D램|DRAM|디램|HBM|DDR5|DDR4|LPDDR|낸드|NAND|메모리|memory chip|메모리반도체")
PRICE_KW = re.compile(r"가격|계약가|고정거래|스팟|spot|ASP|price|인상|인하|하락|급등|급락|품귀|부족|과잉|반등|상승|약세|강세|DXI|수요|증설")

UP = re.compile(r"인상|급등|상승|반등|오른|오름|품귀|부족|2배|강세|치솟|뛰|최고가|상향|호황|훈풍|초과수요|공급부족")
DOWN = re.compile(r"인하|하락|급락|약세|하락세|내림|하향|떨어|저가|최저|공급과잉|둔화|부진|재고증가|가격 압박")


def _fetch_rss(query):
    url = ("https://news.google.com/rss/search?q=" + urllib.parse.quote(query) +
           "&hl=ko&gl=KR&ceid=KR:ko")
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    raw = urllib.request.urlopen(req, timeout=12).read().decode("utf-8", "replace")
    out = []
    for it in re.findall(r"<item>(.*?)</item>", raw, re.S):
        t = re.search(r"<title>(.*?)</title>", it, re.S)
        l = re.search(r"<link>(.*?)</link>", it, re.S)
        d = re.search(r"<pubDate>(.*?)</pubDate>", it)
        s = re.search(r"<source[^>]*>(.*?)</source>", it, re.S)
        if not t:
            continue
        title = _html.unescape(re.sub(r"<[^>]+>", "", t.group(1))).strip()
        # Google은 제목 끝에 ' - 언론사' 부착 → 분리
        src = _html.unescape(s.group(1)).strip() if s else ""
        title = re.sub(r"\s*-\s*" + re.escape(src) + r"\s*$", "", title) if src else title
        try:
            dt = parsedate_to_datetime(d.group(1)).astimezone(KST) if d else None
        except Exception:
            dt = None
        out.append({"title": title, "url": (l.group(1).strip() if l else ""),
                    "source": src, "dt": dt})
    return out


def _tag(title):
    if UP.search(title) and not DOWN.search(title):
        return "up"
    if DOWN.search(title) and not UP.search(title):
        return "down"
    if UP.search(title) and DOWN.search(title):
        return "mixed"
    return "neutral"


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    seen_titles = set()
    articles = []
    for q in QUERIES:
        try:
            items = _fetch_rss(q)
        except Exception as e:
            print(f"  [WARN] RSS 실패({q[:20]}): {e}")
            continue
        for it in items:
            title = it["title"]
            key = re.sub(r"\s+", "", title)[:40]
            if key in seen_titles:
                continue
            if not (MEM_KW.search(title) and PRICE_KW.search(title)):
                continue  # 메모리+가격 둘 다 있어야(노이즈 제거)
            seen_titles.add(key)
            it["direction"] = _tag(title)
            articles.append(it)

    # 최신순 (날짜 없으면 뒤로)
    articles.sort(key=lambda a: a["dt"] or datetime(1970, 1, 1, tzinfo=KST), reverse=True)
    articles = articles[:24]

    up_n = sum(1 for a in articles if a["direction"] == "up")
    down_n = sum(1 for a in articles if a["direction"] == "down")
    # 최근 10건 기준 가격 모멘텀
    recent = articles[:10]
    ru = sum(1 for a in recent if a["direction"] == "up")
    rd = sum(1 for a in recent if a["direction"] == "down")
    if ru - rd >= 2:
        momentum, mcol = "🟢 가격 상승 우위 (메모리 수혜)", "green"
    elif rd - ru >= 2:
        momentum, mcol = "🔴 가격 하락 우위 (메모리 부담)", "red"
    else:
        momentum, mcol = "⚪ 혼조·중립", "gray"

    out_articles = [{
        "title": a["title"], "url": a["url"], "source": a["source"],
        "date": a["dt"].strftime("%Y-%m-%d") if a["dt"] else "",
        "direction": a["direction"],
    } for a in articles]

    # 신규 가격방향 뉴스 텔레그램 경보
    new_alerts = []
    prev_seen = set()
    try:
        from core import load_state
        prev_seen = set((load_state("memory_price", default={}) or {}).get("seen", []))
    except Exception:
        pass
    for a in out_articles:
        key = re.sub(r"\s+", "", a["title"])[:40]
        if key not in prev_seen and a["direction"] in ("up", "down"):
            new_alerts.append(a)

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "momentum": momentum, "momentum_color": mcol,
        "up_count": up_n, "down_count": down_n, "n": len(out_articles),
        "articles": out_articles,
        "note": ("메모리 가격 '뉴스' 자동 추적(Google News). DRAM 스팟가·DXI·계약가 원본은 TrendForce 유료. "
                 "HBM은 계약 기반이라 실시간 가격 없음 — 뉴스로만 추적. 가격↑=SK·삼성 수혜. 정보용·투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"메모리 가격 뉴스 {len(out_articles)}건 · 상승 {up_n} 하락 {down_n} · {momentum}")
    for a in out_articles[:6]:
        arrow = {"up": "▲", "down": "▼", "mixed": "↕", "neutral": "·"}[a["direction"]]
        print(f"  {arrow} [{a['date']}] {a['title'][:50]}")

    if new_alerts:
        try:
            from core import send_message, get_secret, save_state
            if get_secret("TELEGRAM_FINANCE_BOT_TOKEN"):
                lines = [f"{'▲인상' if a['direction']=='up' else '▼하락'} {a['title'][:56]}" for a in new_alerts[:6]]
                body = "💾 메모리 가격 뉴스 (신규)\n" + "\n".join("• " + x for x in lines)
                if len(new_alerts) > 6:
                    body += f"\n…외 {len(new_alerts)-6}건"
                if send_message(body):
                    print(f"  📨 메모리 가격 경보 {len(new_alerts)}건 발송")
        except Exception as e:
            print(f"  [WARN] 텔레그램 실패: {e}")

    try:
        from core import save_state
        save_state("memory_price", {"seen": [re.sub(r"\s+", "", a["title"])[:40] for a in out_articles][:200]})
    except Exception:
        pass

    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
