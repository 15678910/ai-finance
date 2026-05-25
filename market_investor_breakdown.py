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
    """단일 시장의 최근 N 거래일 데이터 수집 (페이지 1개로 충분 — 30일 표시됨)."""
    today_str = datetime.now().strftime("%Y%m%d")
    by_date = fetch_breakdown_for_date(market_code, today_str)
    if not by_date:
        return []
    sorted_rows = sorted(by_date.values(), key=lambda r: r["date"], reverse=True)
    return sorted_rows[:days_target]


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

    # JSON 저장
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "market_investor_breakdown.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f">>> 저장 완료: {out_path}")


if __name__ == "__main__":
    main()
