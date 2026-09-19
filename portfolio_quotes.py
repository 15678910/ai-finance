"""
넓은 시세 파일 — 코스피+코스닥 시가총액 상위 종목 현재가 (포트폴리오 자동 매칭용)
=================================================================================
포트폴리오 손익 패널이 '주요 130종목' 밖 종목도 자동 매칭하도록 코드·현재가를
폭넓게 수집(코스피+코스닥 각 상위 ~750종목). 가격만 담으므로 공개(개인정보 아님).
원금·보유종목은 브라우저(localStorage)에만 저장.

출력: docs/quotes.json  { "generated_at":..., "count":..., "source":...,
                          "q": { "005930": {"n":"삼성전자","p":309500}, ... } }

수집 경로 — 순서대로 시도, 앞이 충분히 성공하면 뒤는 생략
  1. 네이버 모바일 시세 API (JSON)  m.stock.naver.com/api/stocks/marketValue/{KOSPI|KOSDAQ}
  2. 네이버 실시간 폴링 API (JSON)  polling.finance.naver.com/api/realtime/domestic/stock/{codes}
     — naver-stock-proxy-worker.js 가 쓰는 엔드포인트. 직전 quotes.json 의 종목코드로
       가격만 다시 받는다(종목 목록은 못 늘리지만 시세 정체는 막는다).
  3. 네이버 PC 시세 페이지 HTML 정규식 — 예전 경로. 2026-09-10부터 0건.

왜 바꿨나 (2026-09-19)
  HTML 정규식이 페이지 구조 변경으로 0건을 돌려주는데, 스텝이 `|| echo` 로 감싸여
  워크플로는 초록색이라 8일 넘게 아무도 몰랐다(프리시니스 배너만 켜짐).
  · JSON API 는 HTML 보다 구조가 안정적이다.
  · 어느 경로든 0건이면 응답 길이·앞부분을 로그에 남긴다 — 다음 장애는 Actions 로그만
    보고 진단할 수 있어야 한다. 이번엔 "0종목" 한 줄뿐이라 원인을 알 수 없었다.
  · 수집이 100종목 미만이면 기존 파일을 보존한다(예전과 동일).

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

MIN_ROWS = 100          # 이보다 적으면 실패로 보고 기존 파일 보존
MOBILE_PAGE_SIZE = 100
MOBILE_MAX_PAGES = 8    # 시장별 ~800종목
HTML_PAGES = 15         # 시장별 페이지 수(1페이지=50종목) → 각 ~750종목
POLL_BATCH = 50         # 폴링 API 한 번에 넘길 종목 수

MOBILE_URL = "https://m.stock.naver.com/api/stocks/marketValue/{market}?page={page}&pageSize={size}"
POLL_URL = "https://polling.finance.naver.com/api/realtime/domestic/stock/{codes}"
HTML_URL = "https://finance.naver.com/sise/sise_market_sum.naver?sosok={sosok}&page={page}"

ROW = re.compile(r'/item/main\.naver\?code=(\d{6})"\s+class="tltle">([^<]+)</a>\s*</td>\s*<td class="number">([\d,]+)</td>')


# ====================================================================
# HTTP · 진단
# ====================================================================
def _get(url, accept="*/*", encoding=None, timeout=12):
    req = urllib.request.Request(url, headers={
        "User-Agent": UA, "Referer": "https://finance.naver.com/", "Accept": accept,
    })
    raw = urllib.request.urlopen(req, timeout=timeout).read()
    return raw.decode(encoding or "utf-8", errors="replace")


def _diag(label, text):
    """0건일 때 응답이 '무엇이었는지' 로그에 남긴다 — 차단 페이지인지, 구조 변경인지."""
    body = re.sub(r"\s+", " ", (text or ""))[:300]
    print(f"  [DIAG] {label}: 길이 {len(text or '')} · 앞부분: {body!r}")


def _to_int(v):
    try:
        return int(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


# ====================================================================
# 파서 — 순수 함수 (테스트 대상)
# ====================================================================
def parse_ranking_json(payload):
    """모바일 시세 API 응답 → {code: {"n": 이름, "p": 가격}}.

    응답 최상위가 리스트일 수도, {"stocks":[...]} 처럼 감싸여 있을 수도 있어
    후보 키를 순서대로 찾는다. 필드명도 itemCode/cd, stockName/nm, closePrice/nowVal 로
    변종이 있어 모두 받는다.
    """
    items = None
    if isinstance(payload, list):
        items = payload
    elif isinstance(payload, dict):
        for k in ("stocks", "result", "datas", "items", "list"):
            v = payload.get(k)
            if isinstance(v, list):
                items = v
                break
            if isinstance(v, dict):
                for kk in ("stocks", "items", "list"):
                    if isinstance(v.get(kk), list):
                        items = v[kk]
                        break
            if items is not None:
                break
    out = {}
    for it in items or []:
        if not isinstance(it, dict):
            continue
        code = str(it.get("itemCode") or it.get("cd") or it.get("code") or "").strip()
        name = it.get("stockName") or it.get("nm") or it.get("name") or ""
        price = _to_int(it.get("closePrice") or it.get("nowVal") or it.get("price") or it.get("nv"))
        if re.fullmatch(r"\d{6}", code) and price and price > 0 and code not in out:
            out[code] = {"n": _html.unescape(str(name)).strip() or code, "p": price}
    return out


def parse_polling_json(payload):
    """실시간 폴링 API 응답 → {code: {"n", "p"}}. 형태는 워커 코드에 문서화된 그대로:
    {"datas":[{"itemCode":"000660","stockName":"SK하이닉스","closePrice":"2,343,000",...}]}"""
    datas = (payload or {}).get("datas") if isinstance(payload, dict) else None
    out = {}
    for d in datas or []:
        if not isinstance(d, dict):
            continue
        code = str(d.get("itemCode") or d.get("cd") or "").strip()
        price = _to_int(d.get("closePrice"))
        if re.fullmatch(r"\d{6}", code) and price and price > 0:
            out[code] = {"n": str(d.get("stockName") or code).strip(), "p": price}
    return out


def parse_html_rows(html_text):
    """PC 시세 페이지 HTML → {code: {"n", "p"}} (예전 정규식)."""
    out = {}
    for code, name, price in ROW.findall(html_text or ""):
        p = _to_int(price)
        if p and p > 0 and code not in out:
            out[code] = {"n": _html.unescape(name).strip(), "p": p}
    return out


# ====================================================================
# 수집 경로
# ====================================================================
def fetch_ranking_api(market):
    """1순위. 페이지를 돌며 새 종목이 안 늘면 멈춘다."""
    out = {}
    for page in range(1, MOBILE_MAX_PAGES + 1):
        url = MOBILE_URL.format(market=market, page=page, size=MOBILE_PAGE_SIZE)
        text = ""
        try:
            text = _get(url, accept="application/json")
            rows = parse_ranking_json(json.loads(text))
        except Exception as e:
            print(f"  [WARN] ranking {market} p{page}: {type(e).__name__}: {e}")
            if page == 1 and text:
                _diag(f"ranking {market} p1", text)   # JSON 파싱 실패면 원문을 남긴다
            break
        if page == 1 and not rows:
            _diag(f"ranking {market} p1 (0건)", text)
        n0 = len(out)
        out.update({k: v for k, v in rows.items() if k not in out})
        if len(out) == n0:
            break
        time.sleep(0.12)
    return out


def load_previous_codes():
    """직전 quotes.json 의 종목코드 — 폴링 API 로 가격만 갱신할 때 쓴다."""
    try:
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            return list((json.load(f).get("q") or {}).keys())
    except Exception:
        return []


def fetch_polling_api(codes):
    """2순위. 코드 목록을 배치로 나눠 가격을 받는다."""
    out = {}
    for i in range(0, len(codes), POLL_BATCH):
        batch = codes[i:i + POLL_BATCH]
        url = POLL_URL.format(codes=",".join(batch))
        try:
            text = _get(url, accept="application/json")
            rows = parse_polling_json(json.loads(text))
        except Exception as e:
            print(f"  [WARN] polling batch {i // POLL_BATCH + 1}: {type(e).__name__}: {e}")
            continue
        if i == 0 and not rows:
            _diag("polling batch1 (0건)", text)
        out.update(rows)
        time.sleep(0.1)
    return out


def scrape_html(sosok):
    """3순위. 예전 경로 — 구조가 바뀌면 0건. 0건이면 무엇을 받았는지 남긴다."""
    out = {}
    for pg in range(1, HTML_PAGES + 1):
        try:
            html_text = _get(HTML_URL.format(sosok=sosok, page=pg), encoding="euc-kr")
        except Exception as e:
            print(f"  [WARN] html sosok={sosok} page={pg}: {e}")
            continue
        rows = parse_html_rows(html_text)
        if pg == 1 and not rows:
            _diag(f"html sosok={sosok} p1 (0건)", html_text)
        n0 = len(out)
        out.update({k: v for k, v in rows.items() if k not in out})
        if len(out) == n0:
            break
        time.sleep(0.12)
    return out


def collect():
    """(quotes, source). 앞 경로가 MIN_ROWS 이상 모으면 거기서 끝낸다."""
    # 1) 모바일 시세 API
    q = {}
    for market in ("KOSPI", "KOSDAQ"):
        part = fetch_ranking_api(market)
        print(f"  ranking-api {market}: {len(part)}종목")
        q.update(part)
    if len(q) >= MIN_ROWS:
        return q, "naver-mobile-api"

    # 2) 폴링 API — 직전 파일의 종목코드로 가격 갱신
    codes = load_previous_codes()
    if codes:
        part = fetch_polling_api(codes)
        print(f"  polling-api: {len(part)}/{len(codes)}종목")
        if len(part) >= MIN_ROWS:
            return part, "naver-polling-api"
    else:
        print("  polling-api: 직전 quotes.json 없음 — 건너뜀")

    # 3) HTML 정규식
    q = {}
    for name, sosok in (("코스피", 0), ("코스닥", 1)):
        part = scrape_html(sosok)
        print(f"  html {name}: {len(part)}종목")
        q.update(part)
    return q, "naver-html"


# ====================================================================
# 메인
# ====================================================================
def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    quotes, source = collect()

    if len(quotes) < MIN_ROWS:
        print(f"[ERROR] 수집 부족({len(quotes)}, 경로={source}) — 기존 파일 보존. "
              f"위 [DIAG] 줄로 응답 형태를 확인할 것.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "count": len(quotes),
        "source": source,
        "q": quotes,
        "note": "코스피+코스닥 시총 상위 현재가(네이버). 포트폴리오 손익 자동 매칭용. 가격만·공개·실시간 아님.",
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE}  총 {len(quotes)}종목 (경로: {source})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
