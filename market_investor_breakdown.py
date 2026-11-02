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
    """단일 시장의 최근 N 거래일. KRX 정보데이터시스템(JSON) → 실패 시 네이버 PC 페이지.
    2026-10-30 이후 네이버 PC 페이지가 Actions 에서 데이터를 돌려주지 않아 이 섹션이 비어 있었다."""
    rows = fetch_krx_breakdown(market_code, days_target)
    if rows:
        return rows[:days_target]
    today_str = datetime.now().strftime("%Y%m%d")
    by_date = fetch_breakdown_for_date(market_code, today_str)
    if not by_date:
        return []
    sorted_rows = sorted(by_date.values(), key=lambda r: r["date"], reverse=True)
    return sorted_rows[:days_target]


# ====================================================================
# KRX 정보데이터시스템 — 투자자별 거래실적 일별추이 (MDCSTAT02203)
# ====================================================================
KRX_URL = "https://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"
KRX_BLD = "dbms/MDC/STAT/standard/MDCSTAT02203"
KRX_MKT = {"01": "STK", "02": "KSQ"}          # 네이버 sosok 코드 → KRX mktId
# KRX 응답 TRDVAL1..12 의 순서 (투자자별 거래실적 화면 열 순서와 동일)
KRX_ORDER = ["financial_inv", "insurance", "trust", "private_fund", "bank", "other_financial", "pension",
             "institutional_total", "other_corp", "personal", "foreign", "other_foreign"]
KRX_HEADERS = {"User-Agent": HEADERS["User-Agent"],
               "Referer": "https://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd?menuId=MDC0201020203",
               "X-Requested-With": "XMLHttpRequest"}
_KRX_DIAG = {"done": False}


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
        r = requests.post(KRX_URL, data=params, headers=KRX_HEADERS, timeout=20)
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
    print(f"  [WARN] KRX {KRX_MKT.get(market_code)}: {reason} → 네이버 PC 페이지 시도")
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
