"""
넓은 시세 파일 — 코스피+코스닥 시가총액 상위 종목 현재가 (포트폴리오 자동 매칭용)
=================================================================================
포트폴리오 손익 패널이 '주요 130종목' 밖 종목도 자동 매칭하도록, 네이버 시가총액
순위 페이지에서 코드·현재가를 폭넓게 수집(코스피+코스닥 각 상위 ~750종목).
가격만 담으므로 공개(개인정보 아님). 원금·보유종목은 브라우저(localStorage)에만 저장.

출력: docs/quotes.json  { "YYYYMMDD..": ..., "q": { "005930": {"n":"삼성전자","p":309500}, ... } }
🚨 시세 참고용(워크플로 갱신, 실시간 아님).
"""

import json
import os
import re
import sys
import html as _html
import time
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "quotes.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
PAGES = 15   # 시장별 페이지 수(1페이지=50종목) → 각 ~750종목

ROW = re.compile(r'/item/main\.naver\?code=(\d{6})"\s+class="tltle">([^<]+)</a>\s*</td>\s*<td class="number">([\d,]+)</td>')


def _fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Referer": "https://finance.naver.com/"})
    return urllib.request.urlopen(req, timeout=12).read().decode("euc-kr", errors="replace")


def scrape(sosok):
    out = {}
    for pg in range(1, PAGES + 1):
        try:
            html = _fetch(f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={sosok}&page={pg}")
        except Exception as e:
            print(f"  [WARN] sosok={sosok} page={pg}: {e}")
            continue
        n0 = len(out)
        for code, name, price in ROW.findall(html):
            try:
                p = int(price.replace(",", ""))
            except ValueError:
                continue
            if p > 0 and code not in out:
                out[code] = {"n": _html.unescape(name).strip(), "p": p}
        if len(out) == n0:      # 더 이상 행이 없으면 조기 종료
            break
        time.sleep(0.12)
    return out


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    quotes = {}
    for name, sosok in (("코스피", 0), ("코스닥", 1)):
        q = scrape(sosok)
        quotes.update(q)
        print(f"  {name}: {len(q)}종목")

    if len(quotes) < 100:
        print(f"[ERROR] 수집 부족({len(quotes)}) — 페이지 구조 변경 의심. 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "count": len(quotes),
        "q": quotes,
        "note": "코스피+코스닥 시총 상위 현재가(네이버). 포트폴리오 손익 자동 매칭용. 가격만·공개·실시간 아님.",
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE}  총 {len(quotes)}종목")
    return 0


if __name__ == "__main__":
    sys.exit(main())
