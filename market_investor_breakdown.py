"""
시장 전체 투자자별 매매동향 (8-class 세분화)
=====================================
KOSPI/KOSDAQ 시장 전체의 일별 투자자 유형별 순매수 데이터.

투자자 분류 (Naver 금융):
  개인 / 외국인 / 기관계 / 기타법인
    └ 기관계 하위:
       금융투자 · 보험 · 투신(사모) · 은행 · 기타금융 · 연기금 · 사모

데이터 소스: Naver 금융 investorDealTrendDay
  https://finance.naver.com/sise/investorDealTrendDay.naver?bizdate=YYYYMMDD&sosok=01 (KOSPI)
  https://finance.naver.com/sise/investorDealTrendDay.naver?bizdate=YYYYMMDD&sosok=02 (KOSDAQ)

출력: docs/market_investor_breakdown.json
"""
import os
import json
import time
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta

NAVER_URL = "https://finance.naver.com/sise/investorDealTrendDay.naver"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# 시장 코드
MARKETS = {
    "kospi": "01",
    "kosdaq": "02",
}

# 한국 공휴일/주말 제외용 (간단)
def _last_trading_dates(n=30):
    """오늘부터 거꾸로 n 거래일(주말 제외) 날짜 리스트 반환."""
    out = []
    d = datetime.now()
    while len(out) < n:
        if d.weekday() < 5:  # 0=월, 4=금
            out.append(d.strftime("%Y%m%d"))
        d -= timedelta(days=1)
    return out


def _parse_int(s):
    """+/- 콤마 포함 문자열 → int."""
    if not s or s == "-":
        return 0
    s = s.replace(",", "").replace("+", "").strip()
    try:
        return int(s)
    except (ValueError, TypeError):
        try:
            return int(float(s))
        except (ValueError, TypeError):
            return 0


def fetch_breakdown_for_date(market_code: str, bizdate: str) -> dict | None:
    """단일 날짜의 시장 전체 투자자별 매매동향 파싱."""
    try:
        r = requests.get(
            NAVER_URL,
            params={"bizdate": bizdate, "sosok": market_code},
            headers=HEADERS,
            timeout=10,
        )
        r.encoding = "euc-kr"
        soup = BeautifulSoup(r.text, "html.parser")

        # Naver 페이지의 메인 테이블 찾기 (class=type_1)
        tables = soup.find_all("table")
        if not tables:
            return None

        # 첫 번째 테이블 = 일별 매매동향
        rows = tables[0].find_all("tr")

        # 헤더 두 줄 스킵 (3줄째부터 데이터)
        # 컬럼 구조 (수정된 Naver 신규 컬럼):
        #   [날짜, 개인, 외국인, 기관계, 국가, 기타법인,
        #    금융투자, 보험, 투신(사모), 은행, 기타금융, 연기금]
        # 단, 실제 구조는 헤더 병합 때문에 다를 수 있음 → 첫 데이터 행에서 동적 파싱

        data_by_date = {}
        for tr in rows:
            cells = [c.get_text(strip=True) for c in tr.find_all("td")]
            if len(cells) < 6:
                continue
            # 첫 셀이 날짜 형식인지 확인
            date_raw = cells[0]
            if "." not in date_raw or len(date_raw) < 8:
                continue
            # 26.05.22 (2자리) 또는 2026.05.22 (4자리) → 2026-05-22
            try:
                parts = date_raw.split(".")
                if len(parts) != 3:
                    continue
                yy = parts[0].strip()
                if len(yy) == 2:
                    # 2자리 연도 → 20YY (KRX 시장은 미래 거래 없으므로 20YY 가정)
                    year = "20" + yy
                elif len(yy) == 4:
                    year = yy
                else:
                    continue
                iso_date = f"{year}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"
            except Exception:
                continue

            # Naver 컬럼 매핑 (실측 기준):
            #   cells[0]=날짜  cells[1]=개인  cells[2]=외국인  cells[3]=기관계
            #   cells[4]=금융투자  cells[5]=보험  cells[6]=투신
            #   cells[7]=기타금융  cells[8]=은행  cells[9]=연기금 등
            #   cells[10]=사모  cells[11]=기타법인
            # (단위: 백만원, 음수=순매도)
            row = {
                "date": iso_date,
                "personal": _parse_int(cells[1]) if len(cells) > 1 else 0,
                "foreign": _parse_int(cells[2]) if len(cells) > 2 else 0,
                "institutional_total": _parse_int(cells[3]) if len(cells) > 3 else 0,
                "financial_inv": _parse_int(cells[4]) if len(cells) > 4 else 0,
                "insurance": _parse_int(cells[5]) if len(cells) > 5 else 0,
                "trust": _parse_int(cells[6]) if len(cells) > 6 else 0,
                "other_financial": _parse_int(cells[7]) if len(cells) > 7 else 0,
                "bank": _parse_int(cells[8]) if len(cells) > 8 else 0,
                "pension": _parse_int(cells[9]) if len(cells) > 9 else 0,
                "private_fund": _parse_int(cells[10]) if len(cells) > 10 else 0,
                "other_corp": _parse_int(cells[11]) if len(cells) > 11 else 0,
                "unit": "백만원",
            }
            data_by_date[iso_date] = row

        return data_by_date

    except Exception as e:
        print(f"  [WARN] {market_code}/{bizdate} 파싱 실패: {e}")
        return None


def collect_market(market_code: str, days_target: int = 30) -> list:
    """단일 시장의 최근 N 거래일. KRX 정보데이터시스템(JSON) → 네이버 모바일 3분류(기관 세부 없음) → 네이버 PC 페이지.
    2026-10-30 이후 네이버 PC 페이지가 Actions 에서 데이터를 돌려주지 않아 이 섹션이 비어 있었다.
    반환 행에 detail_available 키가 False 이면 기관 세부(금융투자·연기금 등)가 0 으로 채워진 대체 데이터다."""
    rows = fetch_krx_breakdown(market_code, days_target)
    if rows:
        return rows[:days_target]
    rows = fetch_naver_mobile_3class(market_code, days_target)
    if rows:
        return rows[:days_target]
    rows = aggregate_from_flow_state(market_code, days_target)
    if rows:
        return rows[:days_target]
    today_str = datetime.now().strftime("%Y%m%d")
    by_date = fetch_breakdown_for_date(market_code, today_str)
    if not by_date:
        return []
    sorted_rows = sorted(by_date.values(), key=lambda r: r["date"], reverse=True)
    return sorted_rows[:days_target]


# ====================================================================
# 네이버 모바일 — 지수 투자자별 매매동향 (개인·외국인·기관 3분류, 대체 경로)
# ====================================================================
NAVER_INDEX = {"01": "KOSPI", "02": "KOSDAQ"}
NAVER_MOBILE_URLS = ("https://m.stock.naver.com/api/index/{idx}/investorTrend?pageSize={n}&page=1",
                     "https://m.stock.naver.com/api/index/{idx}/trend?pageSize={n}&page=1",
                     "https://api.stock.naver.com/index/{idx}/trend?pageSize={n}&page=1")
_MOBILE_DIAG = {"done": False}
_ZERO_DETAIL = {k: 0 for k in ("financial_inv", "insurance", "trust", "other_financial", "bank", "pension", "private_fund")}


def _mnum(v):
    try:
        return float(str(v).replace(",", "").replace("+", "").replace("%", "").strip())
    except (TypeError, ValueError):
        return None


def parse_naver_3class(payload) -> list:
    """모바일 지수 투자자 동향 JSON → 3분류 행(백만원). 단위는 값 크기로 추정(억원이면 ×100).
    개인/외국인/기관 키를 못 찾으면 []."""
    import re as _re
    items = payload
    if isinstance(payload, dict):
        for k in ("result", "datas", "items", "list", "trend", "trends"):
            if isinstance(payload.get(k), list):
                items = payload[k]
                break
    if not isinstance(items, list):
        return []
    rows = []
    for it in items:
        if not isinstance(it, dict):
            continue
        d = _re.sub(r"[^0-9]", "", str(it.get("bizdate") or it.get("date") or it.get("localTradedAt") or ""))
        if len(d) < 8:
            continue
        pk = next((k for k in it if _re.match(r"(individual|personal|private)[A-Za-z]*(PureBuy|NetBuy|netBuy|Net|Amount|Value)", k)), None)
        fk = next((k for k in it if _re.match(r"(foreigner|foreign|frgn)[A-Za-z]*(PureBuy|NetBuy|netBuy|Net|Amount|Value)", k)), None)
        ik = next((k for k in it if _re.match(r"(organ|institution|inst)[A-Za-z]*(PureBuy|NetBuy|netBuy|Net|Amount|Value)", k)), None)
        if not (pk and fk and ik):
            continue
        p, f, i = _mnum(it.get(pk)), _mnum(it.get(fk)), _mnum(it.get(ik))
        if p is None or f is None or i is None:
            continue
        rows.append({"date": f"{d[:4]}-{d[4:6]}-{d[6:8]}", "personal": p, "foreign": f, "institutional_total": i})
    if not rows:
        return []
    # 단위 추정: 코스피 외국인 일별 순매수는 보통 수백~수천억 → 억원이면 |값| 10~50,000, 백만원이면 그 100배
    med = sorted(abs(r["foreign"]) for r in rows)[len(rows) // 2]
    scale = 1 if med > 50_000 else 100 if med > 5 else None
    if scale is None:
        return []
    out = []
    for r in rows:
        row = {"date": r["date"], "unit": "백만원", "personal": int(round(r["personal"] * scale)),
               "foreign": int(round(r["foreign"] * scale)), "institutional_total": int(round(r["institutional_total"] * scale)),
               "other_corp": int(round(-(r["personal"] + r["foreign"] + r["institutional_total"]) * scale)),   # 잔여 = 기타법인(추정)
               "detail_available": False, **_ZERO_DETAIL}
        out.append(row)
    out.sort(key=lambda r: r["date"], reverse=True)
    return out


def fetch_naver_mobile_3class(market_code: str, days: int = 30) -> list:
    idx = NAVER_INDEX.get(market_code, "KOSPI")
    for url in NAVER_MOBILE_URLS:
        u = url.format(idx=idx, n=days)
        text = ""
        try:
            r = requests.get(u, headers={"User-Agent": HEADERS["User-Agent"], "Referer": "https://m.stock.naver.com/", "Accept": "application/json"}, timeout=15)
            text = r.text
            rows = parse_naver_3class(r.json())
            if rows:
                print(f"    네이버 모바일 {idx}: {len(rows)}일 (3분류 — 기관 세부 없음)")
                return rows
            err = "0행"
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
        print(f"  [WARN] 네이버 모바일 {u.split('/')[2]} {u.split('/')[-1][:20]}: {err}")
        print(f"  [DIAG] 응답 앞부분: {text[:300]!r}")
    return []


# ====================================================================
# 자체 수집분 합산 — 수급 스크리너가 저장한 시총 상위 800종목 외국인·기관 순매수 (3분류, 최후 수단 직전)
# ====================================================================
FLOW_STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs", "flow_streak_state.json")
FLOW_MARKET = {"01": "KOSPI", "02": "KOSDAQ"}


def aggregate_flow_state(state: dict, market: str, days: int = 30) -> list:
    """{market:{code:시장}, flows:{code:[[date,f_shares,i_shares,close],...]}} → 일별 합산 행(백만원).
    개인 = -(외국인+기관) 잔여(기타법인 포함), 기관 세부 0, detail_available False."""
    flows, mk = (state or {}).get("flows") or {}, (state or {}).get("market") or {}
    by_date = {}
    n_codes = 0
    for code, rows in flows.items():
        if mk.get(code) != market:
            continue
        n_codes += 1
        for r in rows:
            try:
                d, f, i, c = r[0], float(r[1]), float(r[2]), float(r[3])
            except (TypeError, ValueError, IndexError):
                continue
            if not d or not c:
                continue
            acc = by_date.setdefault(d, [0.0, 0.0])
            acc[0] += f * c
            acc[1] += i * c
    if n_codes < 50 or not by_date:
        return []
    out = []
    for d in sorted(by_date, reverse=True)[:days]:
        f_m, i_m = int(round(by_date[d][0] / 1e6)), int(round(by_date[d][1] / 1e6))
        out.append({"date": d, "unit": "백만원", "foreign": f_m, "institutional_total": i_m, "personal": -(f_m + i_m), "other_corp": 0,
                    "detail_available": False, "source_note": f"시총 상위 {n_codes}종목 외국인·기관 순매수 합산(개인=잔여) · 기관 세부 미제공", **_ZERO_DETAIL})
    return out


def aggregate_from_flow_state(market_code: str, days: int = 30) -> list:
    try:
        with open(FLOW_STATE_FILE, encoding="utf-8") as f:
            state = json.load(f)
    except Exception as e:
        print(f"  [WARN] flow_streak 상태 없음: {type(e).__name__}: {e}")
        return []
    rows = aggregate_flow_state(state, FLOW_MARKET.get(market_code, "KOSPI"), days)
    if rows:
        print(f"    자체 합산 {FLOW_MARKET.get(market_code)}: {len(rows)}일 (스크리너 수집분, 기준 {state.get('as_of')})")
    return rows


# ====================================================================
# KRX 정보데이터시스템 — 투자자별 거래실적 일별추이 (MDCSTAT02203)
# ====================================================================
KRX_URL = "https://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"
KRX_PAGE = "https://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd?menuId=MDC0201020203"
KRX_BLD = "dbms/MDC/STAT/standard/MDCSTAT02203"
KRX_MKT = {"01": "STK", "02": "KSQ"}          # 네이버 sosok 코드 → KRX mktId
# KRX 응답 TRDVAL1..12 의 순서 (투자자별 거래실적 화면 열 순서와 동일)
KRX_ORDER = ["financial_inv", "insurance", "trust", "private_fund", "bank", "other_financial", "pension",
             "institutional_total", "other_corp", "personal", "foreign", "other_foreign"]
KRX_HEADERS = {"User-Agent": HEADERS["User-Agent"], "Referer": KRX_PAGE, "Origin": "https://data.krx.co.kr",
               "X-Requested-With": "XMLHttpRequest", "Accept": "application/json, text/javascript, */*; q=0.01"}
_KRX_DIAG = {"done": False}
_KRX_SESSION = {"s": None}


def _krx_session():
    """첫 호출 때 메뉴 페이지를 GET 해 세션 쿠키를 받는다 — 쿠키 없이 POST 하면 본문이 'LOGOUT' 으로 온다(2026-11-02)."""
    if _KRX_SESSION["s"] is None:
        s = requests.Session()
        try:
            s.get(KRX_PAGE, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=15)
        except Exception as e:
            print(f"  [WARN] KRX 세션 페이지: {type(e).__name__}: {e}")
        _KRX_SESSION["s"] = s
    return _KRX_SESSION["s"]


def parse_krx_output(payload) -> list:
    """KRX JSON → 네이버 파서와 같은 행 형식(백만원, 날짜 내림차순). 열 이름이 바뀌면 [] 를 돌려준다."""
    out_rows = (payload or {}).get("output") if isinstance(payload, dict) else None
    if not isinstance(out_rows, list):
        return []
    rows = []
    for it in out_rows:
        if not isinstance(it, dict):
            continue
        d = str(it.get("TRD_DD") or it.get("trdDd") or "").replace("/", "-").replace(".", "-").strip()
        if len(d) != 10:
            continue
        vals = {}
        for idx, key in enumerate(KRX_ORDER, start=1):
            v = it.get(f"TRDVAL{idx}")
            if v is None:
                break
            vals[key] = _parse_int(str(v).replace("원", ""))
        if len(vals) < 11:
            continue
        # 원 단위 → 백만원 (KRX 는 '원' 단위를 쓴다; 값이 작으면 이미 백만원일 수 있어 자릿수로 판단)
        scale = 1_000_000 if max(abs(x) for x in vals.values()) > 10_000_000_000 else 1
        row = {"date": d, "unit": "백만원"}
        for key in KRX_ORDER:
            if key == "other_foreign":
                continue
            row[key] = int(round(vals[key] / scale))
        # 네이버 화면의 '외국인'은 등록외국인+기타외국인 합계
        row["foreign"] = int(round((vals["foreign"] + vals.get("other_foreign", 0)) / scale))
        rows.append(row)
    rows.sort(key=lambda r: r["date"], reverse=True)
    return rows


def validate_breakdown(rows: list) -> bool:
    """순매수는 투자자 전체 합이 0 에 가까워야 한다(누군가 사면 누군가 판다). 열 매핑이 틀리면 크게 어긋난다."""
    if not rows:
        return False
    bad = 0
    for r in rows[:10]:
        parts = [r.get("personal", 0), r.get("foreign", 0), r.get("institutional_total", 0), r.get("other_corp", 0)]
        tot = sum(abs(x) for x in parts)
        if tot and abs(sum(parts)) / tot > 0.05:
            bad += 1
    return bad <= 2


def fetch_krx_breakdown(market_code: str, days: int = 30) -> list:
    """KRX 투자자별 거래실적(순매수·거래대금) 최근 N 거래일. 실패·검증 실패 시 []."""
    end = datetime.now()
    start = end - timedelta(days=int(days * 1.8) + 7)
    params = {"bld": KRX_BLD, "locale": "ko_KR", "inqTpCd": "2", "trdVolVal": "2", "askBid": "3",
              "mktId": KRX_MKT.get(market_code, "STK"), "strtDd": start.strftime("%Y%m%d"), "endDd": end.strftime("%Y%m%d"),
              "share": "1", "money": "1", "csvxls_isNo": "false"}
    text = ""
    try:
        r = _krx_session().post(KRX_URL, data=params, headers=KRX_HEADERS, timeout=20)
        text = r.text
        payload = r.json()
        rows = parse_krx_output(payload)
        if rows and validate_breakdown(rows):
            print(f"    KRX {KRX_MKT.get(market_code)}: {len(rows)}일 (최근 {rows[0]['date']})")
            return rows
        reason = "0행 또는 열 매핑 검증 실패"
        if not _KRX_DIAG["done"]:
            _KRX_DIAG["done"] = True
            first = (payload.get("output") or [{}])[0] if isinstance(payload, dict) and isinstance(payload.get("output"), list) else payload
            print(f"  [DIAG] KRX 응답 첫 행: {str(first)[:400]}")
    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        if not _KRX_DIAG["done"]:
            _KRX_DIAG["done"] = True
            print(f"  [DIAG] KRX 응답 앞부분: {text[:300]!r}")
    print(f"  [WARN] KRX {KRX_MKT.get(market_code)}: {reason} → 네이버 모바일/PC 페이지 시도")
    return []


def summarize_recent(rows: list, n: int = 5) -> dict:
    """최근 n 거래일 누적 합계."""
    if not rows:
        return {}
    take = rows[:n]
    keys = ["personal", "foreign", "institutional_total", "financial_inv",
            "insurance", "trust", "other_financial", "bank",
            "pension", "private_fund", "other_corp"]
    return {k: sum(r.get(k, 0) for r in take) for k in keys}


def main():
    print(">>> 시장 투자자 매매동향 (8-class 세분화) 수집 시작")

    output = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S KST"),
        "data_source": "Naver 금융 investorDealTrendDay",
        "unit_note": "단위: 백만원 (음수=순매도, 양수=순매수)",
        "investor_types": {
            "personal": "개인",
            "foreign": "외국인",
            "institutional_total": "기관계",
            "financial_inv": "금융투자",
            "insurance": "보험",
            "trust": "투신",
            "other_financial": "기타금융",
            "bank": "은행",
            "pension": "연기금",
            "private_fund": "사모",
            "other_corp": "기타법인",
        },
        "markets": {},
        "warning": "시뮬레이션/참고용. 정확도 검증 필요.",
    }

    for market_name, code in MARKETS.items():
        print(f"  > {market_name.upper()} 수집...")
        rows = collect_market(code, days_target=30)
        if rows:
            output["markets"][market_name] = {
                "rows": rows,
                "days": len(rows),
                "summary_5d": summarize_recent(rows, 5),
                "summary_20d": summarize_recent(rows, 20),
                "latest_date": rows[0]["date"] if rows else None,
                "detail_available": rows[0].get("detail_available", True),
                "source_note": rows[0].get("source_note"),
            }
            print(f"    수집 완료: {len(rows)}일, 최근일={rows[0]['date']}")
            # 5일 누적 출력
            s5 = output["markets"][market_name]["summary_5d"]
            print(f"    [5일 누적] 개인 {s5['personal']:+,} · 외국인 {s5['foreign']:+,} · 기관계 {s5['institutional_total']:+,}")
            print(f"               연기금 {s5['pension']:+,} · 금융투자 {s5['financial_inv']:+,} · 투신 {s5['trust']:+,}")
        else:
            output["markets"][market_name] = {"rows": [], "error": "수집 실패"}
            print(f"    [FAIL] 수집 실패")
        time.sleep(0.3)

    # JSON 저장 — 전부 실패했으면 직전 파일을 덮어쓰지 않고(stale 표시만) 실패 코드로 끝낸다
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "market_investor_breakdown.json")
    ok_markets = [m for m, v in output["markets"].items() if v.get("rows")]
    if not ok_markets:
        try:
            with open(out_path, encoding="utf-8") as f:
                prev = json.load(f)
            prev["stale"] = True
            prev["stale_note"] = f"{output['generated_at']} 수집 실패 — 직전 데이터 유지"
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(prev, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        print(">>> [FAIL] 두 시장 모두 수집 실패 — 직전 데이터 유지")
        return 1
    output["data_source"] = "KRX 정보데이터시스템 (투자자별 거래실적) · 대체: Naver 금융 investorDealTrendDay"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f">>> 저장 완료: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
