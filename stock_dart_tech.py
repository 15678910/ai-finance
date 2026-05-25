"""
기업 기술 정보 수집 (DART OpenAPI + KIPRIS)
===========================================
DART OpenAPI로 R&D 투자액·기술 공시, KIPRIS로 특허 현황 수집.

환경변수:
  DART_API_KEY   — opendart.fss.or.kr 무료 등록 후 발급
  KIPRIS_API_KEY — plus.kipris.or.kr  무료 등록 후 발급 (선택)

출력: docs/stock_tech_info.json
"""

import io
import json
import os
import time
import zipfile
from datetime import datetime, timedelta

import requests
from bs4 import BeautifulSoup

DART_API_KEY   = os.environ.get("DART_API_KEY", "")
KIPRIS_API_KEY = os.environ.get("KIPRIS_API_KEY", "")

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# ── 대상 종목 (23종목) ──────────────────────────────────────────────────────
TARGET = [
    # 반도체
    {"ticker": "005930", "name": "삼성전자",          "dart_name": "삼성전자",          "sector": "반도체"},
    {"ticker": "000660", "name": "SK하이닉스",         "dart_name": "SK하이닉스",        "sector": "반도체"},
    {"ticker": "042700", "name": "한미반도체",          "dart_name": "한미반도체",         "sector": "반도체"},
    # 배터리
    {"ticker": "006400", "name": "삼성SDI",            "dart_name": "삼성SDI",           "sector": "배터리"},
    {"ticker": "373220", "name": "LG에너지솔루션",     "dart_name": "엘지에너지솔루션",   "sector": "배터리"},
    # 배터리소재
    {"ticker": "051910", "name": "LG화학",             "dart_name": "LG화학",            "sector": "배터리소재"},
    {"ticker": "247540", "name": "에코프로비엠",        "dart_name": "에코프로비엠",       "sector": "배터리소재"},
    {"ticker": "003670", "name": "포스코퓨처엠",        "dart_name": "포스코퓨처엠",       "sector": "배터리소재"},
    {"ticker": "005490", "name": "POSCO홀딩스",         "dart_name": "POSCO홀딩스",       "sector": "배터리소재"},
    # 전자/부품
    {"ticker": "066570", "name": "LG전자",             "dart_name": "LG전자",            "sector": "전자/부품"},
    {"ticker": "034220", "name": "LG디스플레이",        "dart_name": "LG디스플레이",       "sector": "전자/부품"},
    {"ticker": "011070", "name": "LG이노텍",            "dart_name": "LG이노텍",          "sector": "전자/부품"},
    # 자동차/EV
    {"ticker": "005380", "name": "현대차",             "dart_name": "현대자동차",         "sector": "자동차/EV"},
    {"ticker": "000270", "name": "기아",               "dart_name": "기아",              "sector": "자동차/EV"},
    {"ticker": "012330", "name": "현대모비스",          "dart_name": "현대모비스",         "sector": "자동차/EV"},
    # IT/통신
    {"ticker": "035420", "name": "NAVER",              "dart_name": "NAVER",             "sector": "IT/통신"},
    {"ticker": "035720", "name": "카카오",             "dart_name": "카카오",             "sector": "IT/통신"},
    {"ticker": "017670", "name": "SK텔레콤",           "dart_name": "SK텔레콤",          "sector": "IT/통신"},
    {"ticker": "030200", "name": "KT",                 "dart_name": "케이티",             "sector": "IT/통신"},
    # 바이오
    {"ticker": "207940", "name": "삼성바이오로직스",   "dart_name": "삼성바이오로직스",   "sector": "바이오"},
    {"ticker": "068270", "name": "셀트리온",           "dart_name": "셀트리온",           "sector": "바이오"},
    # 방산/우주
    {"ticker": "012450", "name": "한화에어로스페이스", "dart_name": "한화에어로스페이스", "sector": "방산/우주"},
    {"ticker": "047810", "name": "한국항공우주",       "dart_name": "한국항공우주산업",   "sector": "방산/우주"},
]

# ── DART API 키 없을 때 사용할 정적 기준값 (공개 연간보고서 기준 추정) ───────
STATIC_TECH = {
    # 반도체
    "005930": {"rd_billion": 280000, "rd_pct": 8.20, "patent_count": 67000, "tech_tags": ["GAA공정", "HBM", "파운드리", "갤럭시AI", "엑시노스"]},
    "000660": {"rd_billion": 58000,  "rd_pct": 9.10, "patent_count": 21000, "tech_tags": ["HBM3E", "DRAM", "NAND", "AI메모리", "CXL"]},
    "042700": {"rd_billion": 1200,   "rd_pct": 5.30, "patent_count": 340,   "tech_tags": ["TC본더", "FC본더", "비전시스템", "반도체장비"]},
    # 배터리
    "006400": {"rd_billion": 65000,  "rd_pct": 8.30, "patent_count": 4700,  "tech_tags": ["전고체배터리", "하이니켈NCA", "46파이원통형", "실리콘음극재", "ESS"]},
    "373220": {"rd_billion": 10000,  "rd_pct": 2.90, "patent_count": 25000, "tech_tags": ["원통형4680", "파우치형", "전고체", "NCMA", "배터리BMS"]},
    # 배터리소재
    "051910": {"rd_billion": 12000,  "rd_pct": 2.80, "patent_count": 20000, "tech_tags": ["배터리소재", "양극재", "CNT도전재", "바이오플라스틱", "전해액"]},
    "247540": {"rd_billion": 800,    "rd_pct": 1.80, "patent_count": 1200,  "tech_tags": ["고니켈양극재", "NCA", "NCMA", "하이니켈", "전구체"]},
    "003670": {"rd_billion": 700,    "rd_pct": 1.20, "patent_count": 800,   "tech_tags": ["천연흑연음극재", "인조흑연", "실리콘음극재", "양극재", "LFP"]},
    "005490": {"rd_billion": 6000,   "rd_pct": 1.00, "patent_count": 15000, "tech_tags": ["리튬추출", "수소환원제철", "배터리소재", "포스코리튬", "탄소중립제철"]},
    # 전자/부품
    "066570": {"rd_billion": 31000,  "rd_pct": 5.20, "patent_count": 60000, "tech_tags": ["OLED TV", "전장부품VS", "냉난방HVAC", "가전AI", "로봇클로이"]},
    "034220": {"rd_billion": 11000,  "rd_pct": 5.80, "patent_count": 30000, "tech_tags": ["OLED", "투명OLED", "폴더블OLED", "LTPO", "차량용디스플레이"]},
    "011070": {"rd_billion": 7000,   "rd_pct": 4.90, "patent_count": 12000, "tech_tags": ["카메라모듈", "FC-BGA기판", "자동차부품", "LiDAR", "반도체기판"]},
    # 자동차/EV
    "005380": {"rd_billion": 45000,  "rd_pct": 3.40, "patent_count": 40000, "tech_tags": ["수소연료전지", "IONIQ플랫폼", "보스턴다이나믹스", "AAM에어택시", "SDV"]},
    "000270": {"rd_billion": 20000,  "rd_pct": 3.10, "patent_count": 15000, "tech_tags": ["E-GMP플랫폼", "PBV목적기반차량", "자율주행", "디지털콕핏", "eAxle"]},
    "012330": {"rd_billion": 12000,  "rd_pct": 3.50, "patent_count": 10000, "tech_tags": ["자율주행센서", "전동화부품", "지능형헤드램프", "xEV파워트레인", "ADAS"]},
    # IT/통신
    "035420": {"rd_billion": 21000,  "rd_pct": 18.5, "patent_count": 3800,  "tech_tags": ["하이퍼클로바X", "생성AI", "검색AI", "클라우드", "로보틱스"]},
    "035720": {"rd_billion": 12000,  "rd_pct": 14.2, "patent_count": 1900,  "tech_tags": ["카카오브레인", "AI챗봇", "금융AI", "엔터테크"]},
    "017670": {"rd_billion": 5000,   "rd_pct": 2.20, "patent_count": 5000,  "tech_tags": ["에이닷AI", "양자암호통신", "5G특화망", "AI반도체사피온", "자율주행통신"]},
    "030200": {"rd_billion": 3000,   "rd_pct": 1.30, "patent_count": 4000,  "tech_tags": ["AIDC데이터센터", "믿음AI", "양자통신", "6G연구", "디지털트윈"]},
    # 바이오
    "207940": {"rd_billion": 3000,   "rd_pct": 3.60, "patent_count": 1500,  "tech_tags": ["바이오의약품CDO", "ADC", "세포유전자치료", "mRNA", "항체의약품"]},
    "068270": {"rd_billion": 4000,   "rd_pct": 8.20, "patent_count": 2000,  "tech_tags": ["바이오시밀러", "항체의약품", "짐펜트라", "트룩시마", "허쥬마"]},
    # 방산/우주
    "012450": {"rd_billion": 4000,   "rd_pct": 5.80, "patent_count": 3000,  "tech_tags": ["우주발사체", "항공엔진", "K9자주포", "누리호엔진", "위성체계"]},
    "047810": {"rd_billion": 2500,   "rd_pct": 7.20, "patent_count": 1500,  "tech_tags": ["KF-21보라매", "수리온헬기", "소형위성", "T-50고등훈련기", "달탐사"]},
}

# ── 설비 투자(CapEx) 정적 기준값 (억원, 연간 유형자산 취득 기준) ──────────────
STATIC_CAPEX = {
    # 반도체
    "005930": {"capex_billion": 530000, "capex_note": "반도체 팹·메모리 생산라인"},
    "000660": {"capex_billion": 180000, "capex_note": "DRAM·NAND 팹 증설"},
    "042700": {"capex_billion": 1200,   "capex_note": "본딩장비 제조 설비"},
    # 배터리
    "006400": {"capex_billion": 38000,  "capex_note": "배터리 팩토리 증설"},
    "373220": {"capex_billion": 70000,  "capex_note": "글로벌 배터리 공장 증설"},
    # 배터리소재
    "051910": {"capex_billion": 40000,  "capex_note": "양극재·소재 생산설비"},
    "247540": {"capex_billion": 15000,  "capex_note": "양극재 캐파 증설"},
    "003670": {"capex_billion": 12000,  "capex_note": "음극재·양극재 공장"},
    "005490": {"capex_billion": 50000,  "capex_note": "리튬·배터리소재 생산"},
    # 전자/부품
    "066570": {"capex_billion": 25000,  "capex_note": "전장·가전 생산라인"},
    "034220": {"capex_billion": 20000,  "capex_note": "OLED 패널 생산설비"},
    "011070": {"capex_billion": 10000,  "capex_note": "기판·모듈 생산설비"},
    # 자동차/EV
    "005380": {"capex_billion": 60000,  "capex_note": "EV 전용공장·수소설비"},
    "000270": {"capex_billion": 30000,  "capex_note": "EV 생산라인·PBV"},
    "012330": {"capex_billion": 20000,  "capex_note": "전동화·ADAS 부품설비"},
    # IT/통신
    "035420": {"capex_billion": 6500,   "capex_note": "데이터센터·클라우드"},
    "035720": {"capex_billion": 2500,   "capex_note": "서버·데이터센터"},
    "017670": {"capex_billion": 25000,  "capex_note": "5G 네트워크·AIDC"},
    "030200": {"capex_billion": 25000,  "capex_note": "5G 망·AI데이터센터"},
    # 바이오
    "207940": {"capex_billion": 20000,  "capex_note": "바이오 생산플랜트(5공장)"},
    "068270": {"capex_billion": 3000,   "capex_note": "바이오시밀러 생산설비"},
    # 방산/우주
    "012450": {"capex_billion": 20000,  "capex_note": "항공엔진·방산 생산설비"},
    "047810": {"capex_billion": 1000,   "capex_note": "항공기·위성 조립설비"},
}

# IPC 기술 분류 코드 → 한국어
IPC_MAP = {
    "H01M": "배터리/전지", "H01L": "반도체", "G06F": "컴퓨팅",
    "G06N": "AI/머신러닝", "H04L": "통신네트워크", "B65G": "물류/운반",
    "A23L": "식품가공", "G06Q": "전자상거래/비즈니스", "H04W": "무선통신",
    "B25J": "로봇", "G16H": "헬스케어IT",
}


# ── DART 기능 ──────────────────────────────────────────────────────────────

def _dart_get(path, params):
    params["crtfc_key"] = DART_API_KEY
    try:
        r = requests.get(f"https://opendart.fss.or.kr/api/{path}",
                         params=params, headers=HEADERS, timeout=15)
        return r.json()
    except Exception as e:
        print(f"  [DART WARN] {path}: {e}")
        return {}


def download_corp_code_map():
    """DART 전체 기업코드 ZIP 다운로드 → {stock_code: corp_code}"""
    try:
        r = requests.get("https://opendart.fss.or.kr/api/corpCode.xml",
                         params={"crtfc_key": DART_API_KEY},
                         headers=HEADERS, timeout=30)
        z = zipfile.ZipFile(io.BytesIO(r.content))
        from xml.etree import ElementTree as ET
        root = ET.fromstring(z.read("CORPCODE.xml"))
        mapping = {}
        for item in root.findall("list"):
            stk  = (item.findtext("stock_code") or "").strip()
            corp = (item.findtext("corp_code")  or "").strip()
            if stk:
                mapping[stk] = corp
        print(f"  [DART] 기업코드 매핑 {len(mapping):,}건 로드")
        return mapping
    except Exception as e:
        print(f"  [DART WARN] corpCode 다운로드 실패: {e}")
        return {}


def get_rd_expense(corp_code, year=2024):
    """DART 재무제표 → R&D 비용 (단위: 백만원)"""
    data = _dart_get("fnlttSinglAcnt.json", {
        "corp_code":  corp_code,
        "bsns_year":  str(year),
        "reprt_code": "11011",   # 사업보고서
        "fs_div":     "CFS",     # 연결재무제표
    })
    if data.get("status") != "000":
        return None, None

    RD_KEYS  = {"연구개발비", "경상연구개발비", "연구및개발비", "연구비"}
    REV_KEYS = {"매출액", "영업수익", "수익(매출액)"}

    rd = rev = None
    for item in data.get("list", []):
        nm  = item.get("account_nm", "")
        raw = (item.get("thstrm_amount") or "").replace(",", "")
        try:
            val = int(raw)
        except Exception:
            continue
        if any(k in nm for k in RD_KEYS) and rd is None:
            rd = val
        if any(k in nm for k in REV_KEYS) and rev is None:
            rev = val

    return rd, rev   # 단위: 백만원


def get_capex(corp_code, year=2024):
    """DART 현금흐름표 → 설비 투자(유형자산 취득) (단위: 백만원)

    fnlttSinglAcntAll.json 로 전체 재무제표를 받아 sj_div='CF' 항목 중
    '유형자산의 취득' 계정을 찾아 반환.  지출이므로 원본 음수 → abs() 처리.
    """
    data = _dart_get("fnlttSinglAcntAll.json", {
        "corp_code":  corp_code,
        "bsns_year":  str(year),
        "reprt_code": "11011",   # 사업보고서
        "fs_div":     "CFS",     # 연결재무제표
    })
    if data.get("status") != "000":
        return None

    CAPEX_KEYS = {"유형자산의 취득", "유형자산 취득", "설비투자", "자본적지출",
                  "유형자산의취득", "유형자산취득"}

    for item in data.get("list", []):
        if item.get("sj_div") != "CF":
            continue
        nm = item.get("account_nm", "")
        if any(k in nm for k in CAPEX_KEYS):
            raw = (item.get("thstrm_amount") or "").replace(",", "").lstrip("-")
            try:
                return abs(int(raw))   # 백만원, 부호 제거
            except Exception:
                continue

    return None


def get_tech_disclosures(corp_code, days=180):
    """DART 공시 목록에서 기술 관련 공시 필터링"""
    end   = datetime.now()
    start = end - timedelta(days=days)
    data  = _dart_get("list.json", {
        "corp_code":  corp_code,
        "bgn_de":     start.strftime("%Y%m%d"),
        "end_de":     end.strftime("%Y%m%d"),
        "page_count": "40",
    })
    TECH_KW = ["특허", "기술", "연구", "개발", "기술이전", "지식재산", "R&D", "발명"]
    out = []
    for item in data.get("list", []):
        title = item.get("report_nm", "")
        if any(kw in title for kw in TECH_KW):
            out.append({
                "date":  item.get("rcept_dt", ""),
                "title": title,
                "url":   f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={item.get('rcept_no','')}",
            })
    return out[:10]


# ── KIPRIS 특허 수집 ──────────────────────────────────────────────────────

def get_patent_count_kipris(company_name):
    """KIPRIS Plus API or 웹 스크래핑으로 특허 건수 조회"""
    # 1) API 방식
    if KIPRIS_API_KEY:
        try:
            url = "http://plus.kipris.or.kr/openapi/rest/patUtiModInfoSearchSevice/getPatUtiModInfoSearch"
            r = requests.get(url, params={
                "ServiceKey": KIPRIS_API_KEY,
                "applicant":  company_name,
                "numOfRows":  "1",
                "pageNo":     "1",
            }, headers=HEADERS, timeout=15)
            from xml.etree import ElementTree as ET
            root = ET.fromstring(r.content)
            total = root.findtext(".//totalCount") or "0"
            return int(total)
        except Exception as e:
            print(f"  [KIPRIS WARN] API 실패: {e}")

    # 2) 웹 스크래핑 방식 (API 키 없을 때)
    try:
        r = requests.get(
            "https://patent.kipris.or.kr/patentsearch/retrieveList.do",
            params={"applicant": company_name, "SearchMode": "A"},
            headers=HEADERS, timeout=10,
        )
        soup = BeautifulSoup(r.text, "html.parser")
        total_el = soup.find(class_="totalNum") or soup.find("span", {"id": "totalCount"})
        if total_el:
            txt = total_el.get_text(strip=True).replace(",", "")
            return int("".join(filter(str.isdigit, txt))) if txt else None
    except Exception:
        pass

    return None


# ── 메인 ─────────────────────────────────────────────────────────────────

def main():
    print("=" * 64)
    print("  기업 기술 정보 수집 (DART + KIPRIS)")
    print(f"  DART API: {'✓ 키 있음' if DART_API_KEY else '✗ 키 없음 — 정적 기준값 사용'}")
    print(f"  KIPRIS  : {'✓ 키 있음' if KIPRIS_API_KEY else '△ 웹 스크래핑 시도'}")
    print("=" * 64)

    # DART corp_code 매핑
    corp_map = {}
    if DART_API_KEY:
        corp_map = download_corp_code_map()

    result = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S KST"),
        "data_sources": {
            "dart": "DART OpenAPI (opendart.fss.or.kr)" if DART_API_KEY else "정적 기준값 (DART API 키 미설정)",
            "kipris": "KIPRIS Plus API" if KIPRIS_API_KEY else "KIPRIS 웹 스크래핑",
        },
        "stocks": {},
    }

    for stk in TARGET:
        ticker = stk["ticker"]
        name   = stk["name"]
        print(f"\n  [{ticker}] {name}")

        static      = STATIC_TECH.get(ticker, {})
        static_capex = STATIC_CAPEX.get(ticker, {})
        entry  = {
            "ticker":        ticker,
            "name":          name,
            "sector":        stk["sector"],
            "rd_billion":    static.get("rd_billion"),       # 억원
            "rd_pct":        static.get("rd_pct"),           # 매출 대비 %
            "capex_billion": static_capex.get("capex_billion"),  # 억원
            "capex_note":    static_capex.get("capex_note", ""),
            "patent_count":  static.get("patent_count"),
            "tech_tags":     static.get("tech_tags", []),
            "disclosures":   [],
            "data_quality":  "static",
        }

        # DART API 수집 시도
        if DART_API_KEY and ticker in corp_map:
            corp_code = corp_map[ticker]
            print(f"    corp_code={corp_code}")

            rd, rev = get_rd_expense(corp_code, year=2024)
            if rd is not None:
                entry["rd_billion"]   = round(rd / 100)     # 백만원 → 억원
                entry["rd_pct"]       = round(rd / rev * 100, 2) if rev else None
                entry["data_quality"] = "dart_api"
                print(f"    R&D: {entry['rd_billion']:,}억원 ({entry['rd_pct']}%)")
            else:
                print(f"    R&D: DART 응답 없음 — 정적값 사용")

            capex_raw = get_capex(corp_code, year=2024)
            if capex_raw is not None:
                entry["capex_billion"] = round(capex_raw / 100)  # 백만원 → 억원
                entry["data_quality"]  = "dart_api"
                print(f"    CapEx: {entry['capex_billion']:,}억원 (유형자산 취득)")
            else:
                print(f"    CapEx: DART 응답 없음 — 정적값 사용")

            disclosures = get_tech_disclosures(corp_code)
            entry["disclosures"] = disclosures
            print(f"    기술 공시: {len(disclosures)}건")
            time.sleep(0.3)

        # KIPRIS 특허 수집
        pat = get_patent_count_kipris(stk["dart_name"])
        if pat is not None:
            entry["patent_count"] = pat
            print(f"    특허: {pat:,}건 (KIPRIS)")
        else:
            print(f"    특허: {entry.get('patent_count', '—')}건 (정적값)")

        result["stocks"][ticker] = entry
        time.sleep(0.2)

    # 저장
    out_dir  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")
    out_path = os.path.join(out_dir, "stock_tech_info.json")
    os.makedirs(out_dir, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n>>> 저장 완료: {out_path}")


if __name__ == "__main__":
    main()
