"""
글로벌 시가총액 추적기 — 국내 합계·코스피·코스닥·암호화폐·S&P 500·나스닥100·미국 대형주 합계
=====================================================================================

카드 7개의 시가총액을 매 실행마다 기록해 추세 차트를 그린다.

측정 방식 (정확 / 추정을 화면에 구분 표시):
  • 코스피·코스닥        네이버 모바일 시세 API 전 종목 시총 합산                 → 정확
  • 국내 합계            코스피 + 코스닥                                          → 정확
  • 암호화폐 전체        CoinGecko /global (total_market_cap.usd)                  → 정확
  • S&P 500              [보정] 500 구성종목 yfinance 시총 합산(일 1회, --calibrate)
                         [장중] 보정 시총 × (현재 지수 / 보정 시점 지수)          → 추정(오차 ~1%)
  • 나스닥100            같은 방식(^NDX). 나스닥 종합(3천 종목)은 무료 소스 없음   → 추정
  • 미국 대형주 합계     S&P 500 ∪ 나스닥100 (중복 제거) — "세계 전체"가 아님       → 추정

이력이 없으면 --calibrate 시 지수 1년 일봉으로 역산 백필(est=True, 차트에 점선).
실패 시 직전 값 유지 + stale 표시, 파싱 불가 응답은 [DIAG] 로 앞부분 기록.

실행:
  python market_cap_tracker.py               # 가벼운 갱신 (시간별, overseas-market.yml)
  python market_cap_tracker.py --calibrate   # 구성종목 합산 보정 + 백필 (일 1회, market-cap-tracker.yml)
출력: docs/market_cap_tracker.json  /  상태: docs/market_cap_tracker_state.json
"""

import csv
import io
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone, timedelta
import urllib.request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
from core.state_store import load_state, save_state

KST = timezone(timedelta(hours=9))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "market_cap_tracker.json")
STATE_NAME = "market_cap_tracker"
UA = ("Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 "
      "(KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1")

NAVER_MV_URL = "https://m.stock.naver.com/api/stocks/marketValue/{market}?page={page}&pageSize=100"
NAVER_MAX_PAGES = 25
COINGECKO_GLOBAL = "https://api.coingecko.com/api/v3/global"
COINGECKO_BTC_CHART = "https://api.coingecko.com/api/v3/coins/bitcoin/market_chart?vs_currency=usd&days=365&interval=daily"
SP500_CSV = "https://raw.githubusercontent.com/datasets/s-and-p-500-companies/main/data/constituents.csv"
SP500_WIKI = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
NDX_WIKI = "https://en.wikipedia.org/wiki/Nasdaq-100"

# 단위: 국내 = 조원(KRW trillion), 해외·암호화폐 = 조달러(USD trillion)
SANITY = {          # (min, max) — 벗어나면 파싱 오류로 간주해 버린다
    "kospi": (500, 30000), "kosdaq": (50, 5000),           # 조원
    "crypto": (0.3, 50), "sp500": (10, 200), "ndx": (5, 150), "us_union": (10, 250),   # 조달러
}
FX_FALLBACK = 1390.0
MIN_COVERAGE = 0.95          # 구성종목 시총 수집률 미달이면 보정 거부(직전 보정 유지)
MAX_POINTS = 900             # 카드별 이력 상한 (최근 7일은 전부, 그 이전은 일 1점)
INDEX_SYMBOL = {"kospi": "^KS11", "kosdaq": "^KQ11", "sp500": "^GSPC", "ndx": "^NDX", "us_union": "^GSPC"}

# 위키·CSV 모두 실패했을 때 쓰는 최후 목록 (2025년 기준 — 보정 시 성공한 목록이 state 에 저장되면 그것을 우선 사용)
NDX_FALLBACK = """AAPL MSFT NVDA AMZN META AVGO GOOGL GOOG TSLA COST NFLX TMUS ASML CSCO AMD PEP LIN AZN ADBE ISRG
QCOM INTU PLTR TXN BKNG AMGN AMAT HON CMCSA ARM PANW ADP GILD VRTX MU ADI SBUX LRCX INTC MELI PDD KLAC MDLZ CTAS
REGN CRWD MAR CEG ORLY CDNS SNPS FTNT ABNB DASH MRVL CSX ADSK PYPL WDAY ROP CHTR NXPI PCAR CPRT AEP MNST ROST PAYX
FANG DDOG TTWO AXON FAST KDP EA BKR ODFL VRSK CCEP EXC XEL IDXX GEHC CTSH TEAM ZS LULU KHC CSGP ANSS DXCM CCEP
ON TTD MCHP WBD GFS BIIB MDB CDW""".split()


# ====================================================================
# HTTP · 진단
# ====================================================================
def _get(url, timeout=15, headers=None):
    h = {"User-Agent": UA, "Accept": "*/*"}
    if headers:
        h.update(headers)
    req = urllib.request.Request(url, headers=h)
    return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", errors="replace")


def _get_json(url, **kw):
    return json.loads(_get(url, **kw))


def _diag(label, text):
    body = re.sub(r"\s+", " ", (text or ""))[:300]
    print(f"  [DIAG] {label}: 길이 {len(text or '')} · 앞부분: {body!r}")


def _num(v):
    """'4,954,163' / 4954163 / '12.3' → float, 아니면 None."""
    try:
        s = str(v).replace(",", "").strip()
        return float(s) if s not in ("", "-", "None") else None
    except (TypeError, ValueError):
        return None


def _in_range(key, v):
    lo, hi = SANITY[key]
    return v is not None and lo <= v <= hi


# ====================================================================
# 파서 — 순수 함수 (테스트 대상)
# ====================================================================
# ETF 브랜드는 반드시 '브랜드 + 공백'(KODEX 200, 파워 K200) — '파워로직스'·'삼성바이오로직스' 같은 일반 종목을 잡지 않도록.
ETF_NAME_RE = re.compile(
    r"^(KODEX|TIGER|KBSTAR|RISE|ACE|SOL|PLUS|HANARO|KOSEF|ARIRANG|KINDEX|TIMEFOLIO|WOORI|BNK|히어로즈|마이티|"
    r"파워|FOCUS|KIWOOM|UNICORN|1Q|DAISHIN343|VITA|ITF|에셋플러스|TRUSTON|KCGI|KoAct|마이다스)\s"
    r"|\bETN\b|ETN\(|\bETF\b"
)
STOCK_TYPE_KEYS = ("stockEndType", "stockType", "endType", "securityType", "itemType", "category")
FUND_TYPES = {"etf", "etn", "elw", "fund", "etfetn", "index_fund"}


def is_etf_like(item):
    """ETF·ETN·ELW 등 펀드성 종목이면 True — 지수 시총에 넣으면 기초자산과 이중 계산된다.
    타입 필드가 명시적으로 펀드류이면 True. 그 외(값이 없거나 모르는 값)는 이름의 운용사 브랜드·ETN 표기로 판별."""
    for k in STOCK_TYPE_KEYS:
        v = item.get(k)
        if v is not None and str(v).strip().lower() in FUND_TYPES:
            return True
    name = str(item.get("stockName") or item.get("nm") or item.get("name") or "")
    return bool(ETF_NAME_RE.search(name))


def parse_naver_mv_page(payload):
    """네이버 모바일 marketValue 페이지 → [(code, 시총 억원)] .

    응답은 {"stocks":[{itemCode, stockName, marketValue:"4,954,163", ...}]} 형태이거나
    최상위 리스트. marketValue 는 억원 단위 문자열. 변종 키(marketCap, mktValue)도 받는다.
    ETF·ETN 등 펀드성 종목은 제외한다(2026-09-29 첫 실행: KOSPI 2,138 '종목'은 ETF·ETN 포함 수치였다).
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
    out = []
    for it in items or []:
        if not isinstance(it, dict):
            continue
        code = str(it.get("itemCode") or it.get("cd") or it.get("code") or "").strip()
        mv = None
        for k in ("marketValue", "marketCap", "mktValue", "marketSum"):
            if it.get(k) is not None:
                mv = _num(it.get(k))
                if mv is not None:
                    break
        if re.fullmatch(r"\d{6}", code) and mv and mv > 0 and not is_etf_like(it):
            out.append((code, mv))
    return out


def describe_types(payload):
    """[DIAG] 용: 페이지 안 종목들의 타입 필드 분포 {필드: {값: 개수}} — 실제 필드명을 로그에서 확인하기 위함."""
    items = payload.get("stocks") if isinstance(payload, dict) else payload
    dist = {}
    for it in items or []:
        if not isinstance(it, dict):
            continue
        for k in STOCK_TYPE_KEYS:
            if k in it:
                dist.setdefault(k, {})
                dist[k][str(it[k])] = dist[k].get(str(it[k]), 0) + 1
    return dist


def sum_market_pages(pages):
    """페이지별 [(code, 억원)] 리스트 → (합계 조원, 종목수). 코드 중복은 1회만."""
    seen = {}
    for rows in pages:
        for code, mv in rows:
            seen.setdefault(code, mv)
    return round(sum(seen.values()) / 10000.0, 2), len(seen)


def parse_coingecko_global(payload):
    """CoinGecko /global → {"cap_t": 조달러, "chg_24h_pct": float|None, "btc_dominance": float|None}."""
    if not isinstance(payload, dict):
        return {}
    d = payload.get("data") if isinstance(payload.get("data"), dict) else payload
    cap = _num(((d or {}).get("total_market_cap") or {}).get("usd"))
    if not cap or cap <= 0:
        return {}
    return {"cap_t": round(cap / 1e12, 4),
            "chg_24h_pct": _num(d.get("market_cap_change_percentage_24h_usd")),
            "btc_dominance": _num(((d.get("market_cap_percentage") or {}).get("btc")))}


def parse_sp500_csv(text):
    """datahub constituents.csv → 티커 목록 (야후 표기: BRK.B → BRK-B)."""
    out = []
    try:
        rows = list(csv.DictReader(io.StringIO(text or "")))
    except Exception:
        return []
    for r in rows:
        sym = (r.get("Symbol") or r.get("symbol") or "").strip()
        if re.fullmatch(r"[A-Z][A-Z0-9.\-]{0,6}", sym):
            out.append(sym.replace(".", "-"))
    return out if 480 <= len(out) <= 520 else []


def parse_wiki_tickers(html_text, expect):
    """위키 구성종목 표에서 티커 셀만 뽑는다. expect=(min,max) 개수 범위를 벗어나면 []."""
    if not html_text:
        return []
    # 티커 셀: <td>AAPL</td> 또는 <td><a ...>AAPL</a></td>
    cands = re.findall(r"<td[^>]*>\s*(?:<a[^>]*>)?\s*([A-Z]{1,5}(?:\.[A-Z])?)\s*(?:</a>)?\s*</td>", html_text)
    seen, out = set(), []
    for c in cands:
        c = c.replace(".", "-")
        if c not in seen:
            seen.add(c)
            out.append(c)
    lo, hi = expect
    return out if lo <= len(out) <= hi else []


# ====================================================================
# 이력 관리 — 순수 함수
# ====================================================================
def thin_series(points, now_iso, max_points=MAX_POINTS):
    """최근 7일은 모든 점, 그 이전은 날짜별 마지막 점만 남긴다. 시간순 정렬·상한 적용."""
    pts = sorted((p for p in points if p.get("t") and p.get("v") is not None), key=lambda p: p["t"])
    cutoff = (datetime.fromisoformat(now_iso) - timedelta(days=7)).isoformat(timespec="minutes")
    by_day, recent = {}, []
    for p in pts:
        if p["t"] >= cutoff:
            recent.append(p)
        else:
            by_day[p["t"][:10]] = p                     # 같은 날이면 마지막 것으로 교체
    out = list(by_day.values()) + recent
    out.sort(key=lambda p: p["t"])
    return out[-max_points:]


def append_point(series, t_iso, value, est=False):
    """같은 분(timestamp) 이면 교체, 아니면 추가."""
    pt = {"t": t_iso, "v": round(float(value), 4)}
    if est:
        pt["est"] = True
    series = [p for p in series if p.get("t") != t_iso]
    series.append(pt)
    return series


def change_1d(series, now_iso):
    """20시간 이상 이전의 가장 최근 점(직전 거래일 마감에 해당) 대비 변화율. 없으면 None."""
    if not series:
        return None
    cur = series[-1]
    ref_cut = (datetime.fromisoformat(now_iso) - timedelta(hours=20)).isoformat(timespec="minutes")
    ref = None
    for p in series[:-1]:
        if p["t"] <= ref_cut and p.get("v"):
            ref = p
    if not ref or not cur.get("v") or ref["v"] == 0:
        return None
    return round((cur["v"] / ref["v"] - 1) * 100, 2)


def backfill_from_index(cap_now, closes, level_now, before_date=None):
    """지수 일봉 closes=[(date, close)] 과 현재 시총·지수로 과거 시총을 비례 역산 (추정).
    before_date(YYYY-MM-DD) 이상의 날짜는 제외 — 오늘의 미완성 봉이 '15:30' 미래 시각으로 들어가
    실측 점을 덮어쓰는 문제(2026-09-29) 방지."""
    if not cap_now or not level_now or level_now <= 0:
        return []
    out = []
    for d, c in closes:
        if before_date and d >= before_date:
            continue
        if c and c > 0:
            out.append({"t": f"{d}T15:30", "v": round(cap_now * c / level_now, 4), "est": True})
    return out


def scale_calibrated(calib, level_now):
    """보정 시총 × (현재 지수 / 보정 시점 지수). 보정 없으면 None."""
    if not calib or not calib.get("cap_t") or not calib.get("index") or not level_now:
        return None
    return round(calib["cap_t"] * level_now / calib["index"], 4)


# ====================================================================
# 수집기
# ====================================================================
def fetch_naver_total(market):
    """KOSPI/KOSDAQ 전 종목 시총 합계(조원). 실패 시 None."""
    pages, seen = [], set()
    for page in range(1, NAVER_MAX_PAGES + 1):
        url = NAVER_MV_URL.format(market=market, page=page)
        text = ""
        try:
            text = _get(url, headers={"Referer": "https://m.stock.naver.com/", "Accept": "application/json"})
            payload = json.loads(text)
            rows = parse_naver_mv_page(payload)
            if page == 1:
                print(f"  [DIAG] naver {market} 타입 필드 분포: {describe_types(payload) or '없음(이름으로 ETF 판별)'}")
        except Exception as e:
            print(f"  [WARN] naver {market} p{page}: {type(e).__name__}: {e}")
            if page == 1:
                _diag(f"naver {market}", text)
            break
        # 첫 실행(2026-09-29)에서 pageSize=100 요청에 99행이 와 '마지막 페이지'로 오판, 99종목만 합산했다.
        # → 행 수가 아니라 '새 종목이 더 없을 때'만 멈춘다.
        new = [r for r in rows if r[0] not in seen]
        if not new:
            if page == 1:
                _diag(f"naver {market} p1 (0건)", text)
            break
        seen.update(r[0] for r in new)
        pages.append(new)
        time.sleep(0.15)
    if not pages:
        return None, 0
    total, n = sum_market_pages(pages)
    if n < 300:
        print(f"  [WARN] naver {market}: {n}종목만 수집 — 페이지네이션 확인 필요")
    key = "kospi" if market.upper() == "KOSPI" else "kosdaq"
    if not _in_range(key, total):
        print(f"  [WARN] naver {market}: 합계 {total}조 sanity 범위 밖 → 버림")
        return None, n
    return total, n


def fetch_crypto():
    try:
        return parse_coingecko_global(_get_json(COINGECKO_GLOBAL))
    except Exception as e:
        print(f"  [WARN] coingecko global: {type(e).__name__}: {e}")
        return {}


def _yf():
    import yfinance as yf
    return yf


def _fast(fi, camel, snake):
    """yfinance FastInfo 값 읽기. 키는 camelCase('lastPrice')이고 속성은 snake_case(last_price)다 —
    2026-09-29 첫 실행에서 fi.get('last_price') 가 항상 None 을 돌려줘 미국 보정이 통째로 빠졌다.
    키 → 속성 순으로 시도하고, 0·NaN·None 은 실패로 본다."""
    for getter in (lambda: fi[camel], lambda: getattr(fi, snake)):
        try:
            v = getter()
            if v is not None and v == v and float(v) > 0:
                return float(v)
        except Exception:
            continue
    return None


def fetch_index_levels(symbols):
    """{심볼: 현재 지수} — yfinance fast_info. 실패한 심볼은 빠진다."""
    out = {}
    try:
        yf = _yf()
    except Exception as e:
        print(f"  [WARN] yfinance import: {e}")
        return out
    for s in symbols:
        try:
            v = _fast(yf.Ticker(s).fast_info, "lastPrice", "last_price")
            if v:
                out[s] = v
            else:
                print(f"  [WARN] index {s}: fast_info 에 lastPrice 없음")
        except Exception as e:
            print(f"  [WARN] index {s}: {type(e).__name__}: {e}")
    return out


def fetch_index_history(symbol, period="1y"):
    """[(YYYY-MM-DD, close)] 일봉."""
    try:
        hist = _yf().Ticker(symbol).history(period=period, interval="1d", auto_adjust=False)
        return [(idx.strftime("%Y-%m-%d"), float(c)) for idx, c in hist["Close"].items() if c == c]
    except Exception as e:
        print(f"  [WARN] history {symbol}: {type(e).__name__}: {e}")
        return []


def fetch_usd_krw(last_fx=None):
    lv = fetch_index_levels(["KRW=X"]).get("KRW=X")
    if lv and 500 <= lv <= 3000:
        return lv, "yfinance KRW=X"
    if last_fx and 500 <= last_fx <= 3000:
        return last_fx, "stale"
    return FX_FALLBACK, "fallback"


def fetch_constituents(state):
    """S&P 500 · 나스닥100 티커 목록. 성공하면 state['lists'] 에 저장, 실패하면 저장분→내장 목록."""
    lists = dict(state.get("lists") or {})
    sp = []
    try:
        sp = parse_sp500_csv(_get(SP500_CSV))
    except Exception as e:
        print(f"  [WARN] sp500 csv: {type(e).__name__}: {e}")
    if not sp:
        try:
            sp = parse_wiki_tickers(_get(SP500_WIKI), (480, 530))
        except Exception as e:
            print(f"  [WARN] sp500 wiki: {type(e).__name__}: {e}")
    if sp:
        lists["sp500"] = sp
    ndx = []
    try:
        ndx = parse_wiki_tickers(_get(NDX_WIKI), (95, 110))
    except Exception as e:
        print(f"  [WARN] ndx wiki: {type(e).__name__}: {e}")
    if ndx:
        lists["ndx"] = ndx
    lists.setdefault("ndx", list(dict.fromkeys(NDX_FALLBACK)))
    print(f"  구성종목: S&P500 {len(lists.get('sp500') or [])} · 나스닥100 {len(lists['ndx'])}")
    return lists


def fetch_market_caps(tickers, workers=8):
    """{티커: 시총 USD} — yfinance fast_info.market_cap, 병렬."""
    try:
        yf = _yf()
    except Exception as e:
        print(f"  [WARN] yfinance import: {e}")
        return {}

    def one(t):
        try:
            return t, _fast(yf.Ticker(t).fast_info, "marketCap", "market_cap"), None
        except Exception as e:
            return t, None, f"{type(e).__name__}: {e}"

    out, first_err = {}, None
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for t, v, err in ex.map(one, tickers):
            if v:
                out[t] = v
            elif err and first_err is None:
                first_err = f"{t}: {err}"
    if len(out) < len(tickers) * 0.5:
        print(f"  [DIAG] 시총 수집 {len(out)}/{len(tickers)} — 첫 오류: {first_err or '오류 없음(값이 None)'}")
    return out


def calibrate(state, levels):
    """구성종목 시총 합산 → state['calib'][key] = {cap_t, index, date, n, n_total}."""
    lists = fetch_constituents(state)
    state["lists"] = lists
    sp, ndx = lists.get("sp500") or [], lists.get("ndx") or []
    union = list(dict.fromkeys(sp + ndx))
    if not union:
        print("  [WARN] 구성종목 목록 없음 → 보정 생략")
        return state
    t0 = time.time()
    caps = fetch_market_caps(union)
    print(f"  시총 수집: {len(caps)}/{len(union)} ({time.time() - t0:.0f}s)")
    calib = dict(state.get("calib") or {})
    today = datetime.now(KST).strftime("%Y-%m-%d")
    for key, members, sym in (("sp500", sp, "^GSPC"), ("ndx", ndx, "^NDX"), ("us_union", union, "^GSPC")):
        if not members:
            continue
        got = [caps[t] for t in members if t in caps]
        cov = len(got) / len(members)
        cap_t = round(sum(got) / 1e12, 4)
        if cov < MIN_COVERAGE:
            print(f"  [WARN] {key}: 수집률 {cov:.0%} < {MIN_COVERAGE:.0%} → 직전 보정 유지")
            continue
        if not _in_range(key, cap_t):
            print(f"  [WARN] {key}: 합계 ${cap_t}T sanity 범위 밖 → 직전 보정 유지")
            continue
        if not levels.get(sym):
            print(f"  [WARN] {key}: 지수 {sym} 없음 → 보정 저장 불가")
            continue
        calib[key] = {"cap_t": cap_t, "index": levels[sym], "date": today, "n": len(got), "n_total": len(members)}
        print(f"  보정 {key}: ${cap_t:.2f}T ({len(got)}/{len(members)}) @ {sym} {levels[sym]:.1f}")
    state["calib"] = calib
    return state


# ====================================================================
# 카드 정의·조립
# ====================================================================
CARDS = [
    # key, 라벨, 그룹, 단위, 기본 품질
    ("domestic", "🇰🇷 국내 합계 (코스피+코스닥)", "국내", "krw_t", "measured"),
    ("kospi", "코스피", "국내", "krw_t", "measured"),
    ("kosdaq", "코스닥", "국내", "krw_t", "measured"),
    ("crypto", "₿ 암호화폐 전체", "암호화폐", "usd_t", "measured"),
    ("sp500", "🇺🇸 S&P 500", "해외", "usd_t", "estimated"),
    ("ndx", "🇺🇸 나스닥100", "해외", "usd_t", "estimated"),
    ("us_union", "🌎 미국 대형주 합계 (S&P500 ∪ 나스닥100)", "해외", "usd_t", "estimated"),
]
CARD_NOTES = {
    "domestic": "네이버 전 종목 시총 합산 · 정확",
    "kospi": "네이버 모바일 API 전 종목 합산 · 정확",
    "kosdaq": "네이버 모바일 API 전 종목 합산 · 정확",
    "crypto": "CoinGecko 전체 코인 시총 · 정확 · 페이지에서 1분마다 실시간 갱신",
    "sp500": "일 1회 500종목 합산 보정 + 장중 지수 비례 추정 (오차 ~1%)",
    "ndx": "나스닥100 기준 (나스닥 종합 3천 종목은 무료 소스 없음) · 지수 비례 추정",
    "us_union": "S&P500과 나스닥100 중복 제거 합계 · '세계 전체'가 아닌 미국 대형주 부분합",
}


def build_output(state, fx, fx_source, now, fresh, quality_override):
    series_all = state.get("series") or {}
    now_iso = now.isoformat(timespec="minutes")
    cards = []
    for key, label, group, unit, base_quality in CARDS:
        ser = series_all.get(key) or []
        cur = ser[-1] if ser else None
        val = cur["v"] if cur else None
        q = quality_override.get(key, base_quality)
        if key not in fresh and val is not None:
            q = "stale"
        if val is None:
            q = "unavailable"
        krw_t = usd_t = None
        if val is not None:
            if unit == "krw_t":
                krw_t, usd_t = val, round(val * 1e12 / fx / 1e12, 4)
            else:
                usd_t, krw_t = val, round(val * fx, 2)
        cards.append({
            "key": key, "label": label, "group": group, "unit": unit,
            "value": val, "value_krw_t": krw_t, "value_usd_t": usd_t,
            "chg_1d_pct": (round(state["crypto_chg_24h"], 2) if key == "crypto" and key in fresh and state.get("crypto_chg_24h") is not None
                           else change_1d(ser, now_iso)),
            "quality": q, "note": CARD_NOTES[key],
            "updated": cur["t"] if cur else None,
            "calib": (state.get("calib") or {}).get(key),
            "n_stocks": (state.get("counts") or {}).get(key),
            "series": [[p["t"], p["v"], 1 if p.get("est") else 0] for p in ser],
        })
    return {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "fx": {"usd_krw": fx, "source": fx_source},
        "cards": cards,
        "note": ("국내·암호화폐는 전 종목 합산(정확), 미국 지수는 일 1회 구성종목 합산 후 장중 지수 비례 추정. "
                 "점선 구간은 지수 일봉으로 역산한 과거 추정치. GitHub Actions 주기(1~2시간)로 갱신되며 암호화폐만 브라우저에서 실시간."),
    }


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    do_calib = "--calibrate" in argv
    now = datetime.now(KST)
    now_iso = now.isoformat(timespec="minutes")
    today = now.strftime("%Y-%m-%d")
    print("=" * 60)
    print(f"  글로벌 시가총액 추적기 {'(보정 모드)' if do_calib else ''}")
    print(f"  KST: {now:%Y-%m-%d %H:%M:%S}")
    print("=" * 60)

    state = load_state(STATE_NAME, {"series": {}, "calib": {}, "lists": {}})
    series_all = state.setdefault("series", {})
    # 스키마 2: 첫 실행(2026-09-29)의 국내 이력은 99종목만 합산된 값과 그 값으로 역산한 백필이라 틀렸다 → 한 번 버리고 다시 쌓는다
    # 스키마 3: 두 번째 실행은 ETF·ETN 이 포함된 합계(KOSPI 2,138 '종목')였다 → 다시 초기화
    if int(state.get("schema") or 1) < 3:
        for k in ("kospi", "kosdaq", "domestic"):
            if series_all.pop(k, None) is not None:
                print(f"  [MIGRATE] {k} 이력 초기화 (ETF·ETN 포함 합산분 제거)")
        state["schema"] = 3
    # 미래 시각 점 제거(오늘 미완성 봉으로 만든 백필) — 실측 점이 마지막이 되도록
    for k in list(series_all):
        series_all[k] = [p for p in series_all[k] if p.get("t", "") <= now_iso]
    fresh = set()
    quality_override = {}

    # ── 환율
    fx, fx_source = fetch_usd_krw(state.get("fx"))
    state["fx"] = fx
    print(f"  USD/KRW {fx:.1f} ({fx_source})")

    # ── 국내
    kospi, n1 = fetch_naver_total("KOSPI")
    kosdaq, n2 = fetch_naver_total("KOSDAQ")
    counts = dict(state.get("counts") or {})
    if kospi:
        series_all["kospi"] = append_point(series_all.get("kospi", []), now_iso, kospi); fresh.add("kospi"); counts["kospi"] = n1
    if kosdaq:
        series_all["kosdaq"] = append_point(series_all.get("kosdaq", []), now_iso, kosdaq); fresh.add("kosdaq"); counts["kosdaq"] = n2
    if kospi and kosdaq:
        series_all["domestic"] = append_point(series_all.get("domestic", []), now_iso, round(kospi + kosdaq, 2)); fresh.add("domestic"); counts["domestic"] = n1 + n2
    state["counts"] = counts
    print(f"  코스피 {kospi if kospi else '실패'}조 ({n1}종목) · 코스닥 {kosdaq if kosdaq else '실패'}조 ({n2}종목)")

    # ── 암호화폐
    cg = fetch_crypto()
    if cg.get("cap_t") and _in_range("crypto", cg["cap_t"]):
        series_all["crypto"] = append_point(series_all.get("crypto", []), now_iso, cg["cap_t"]); fresh.add("crypto")
        state["crypto_chg_24h"] = cg.get("chg_24h_pct")
        state["btc_dominance"] = cg.get("btc_dominance")
    print(f"  암호화폐 ${cg.get('cap_t', '실패')}T")

    # ── 미국 지수
    levels = fetch_index_levels(["^GSPC", "^NDX"])
    if do_calib:
        state = calibrate(state, levels)
    calib = state.get("calib") or {}
    for key in ("sp500", "ndx", "us_union"):
        lv = levels.get(INDEX_SYMBOL[key])
        v = scale_calibrated(calib.get(key), lv)
        if v and _in_range(key, v):
            series_all[key] = append_point(series_all.get(key, []), now_iso, v); fresh.add(key)
            quality_override[key] = "estimated"
    us_txt = " · ".join(f"{k} ${series_all[k][-1]['v']:.2f}T" for k in ("sp500", "ndx", "us_union") if series_all.get(k))
    print("  미국: " + (us_txt or "보정 없음 (--calibrate 필요)"))

    # ── 백필 (보정 모드, 이력이 거의 없는 카드만) — 지수 일봉 비례 역산, est 표시
    if do_calib:
        hist_cache = {}
        for key in ("kospi", "kosdaq", "sp500", "ndx", "us_union"):
            ser = series_all.get(key) or []
            real = [p for p in ser if not p.get("est")]
            if len(real) >= 30 or not ser or key not in fresh:
                continue
            sym = INDEX_SYMBOL[key]
            if sym not in hist_cache:
                hist_cache[sym] = fetch_index_history(sym)
            closes = hist_cache[sym]
            lv = levels.get(sym) or (closes[-1][1] if closes else None)
            if sym in ("^KS11", "^KQ11") and not levels.get(sym):
                lv = fetch_index_levels([sym]).get(sym) or lv
            bf = backfill_from_index(ser[-1]["v"], closes, lv, before_date=today)
            have = {p["t"][:10] for p in ser}
            merged = [p for p in bf if p["t"][:10] not in have] + ser
            series_all[key] = sorted(merged, key=lambda p: p["t"])
            print(f"  백필 {key}: {len(bf)}점 (지수 {sym} 비례 · 추정)")
        # 국내 합계 = 코스피+코스닥 백필 합 (양쪽 같은 날짜만)
        if series_all.get("kospi") and series_all.get("kosdaq") and len([p for p in series_all.get("domestic", []) if not p.get("est")]) < 30:
            kq = {p["t"][:10]: p["v"] for p in series_all["kosdaq"] if p.get("est")}
            have = {p["t"][:10] for p in series_all.get("domestic", [])}
            add = [{"t": p["t"], "v": round(p["v"] + kq[p["t"][:10]], 2), "est": True}
                   for p in series_all["kospi"] if p.get("est") and p["t"][:10] in kq and p["t"][:10] not in have]
            series_all["domestic"] = sorted(add + series_all.get("domestic", []), key=lambda p: p["t"])
        # 암호화폐: BTC 시총 일봉 × (전체/BTC 현재 비율) — 도미넌스 고정 가정, 추정
        if series_all.get("crypto") and len([p for p in series_all["crypto"] if not p.get("est")]) < 30 and state.get("btc_dominance"):
            try:
                chart = _get_json(COINGECKO_BTC_CHART)
                dom = state["btc_dominance"] / 100.0
                have = {p["t"][:10] for p in series_all["crypto"]}
                add = []
                for ts, cap in chart.get("market_caps", []):
                    d = datetime.fromtimestamp(ts / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
                    if d not in have and cap and d < today:
                        add.append({"t": f"{d}T09:00", "v": round(cap / dom / 1e12, 4), "est": True}); have.add(d)
                series_all["crypto"] = sorted(add + series_all["crypto"], key=lambda p: p["t"])
                print(f"  백필 crypto: {len(add)}점 (BTC 시총/도미넌스 · 추정)")
            except Exception as e:
                print(f"  [WARN] crypto backfill: {type(e).__name__}: {e}")

    # ── 이력 정리·저장
    for key in list(series_all):
        series_all[key] = thin_series(series_all[key], now_iso)
    state["last_run"] = now_iso
    save_state(STATE_NAME, state)

    out = build_output(state, fx, fx_source, now, fresh, quality_override)
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    stale = [c["key"] for c in out["cards"] if c["quality"] in ("stale", "unavailable")]
    print(f"[OK] {OUTPUT_FILE}  갱신 {sorted(fresh)} · stale/없음 {stale or '없음'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
