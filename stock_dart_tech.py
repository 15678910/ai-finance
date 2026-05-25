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

# ── 대상 종목 (기존 11종목 + 삼성SDI 추가) ─────────────────────────────────
TARGET = [
    {"ticker": "139480", "name": "이마트",      "dart_name": "이마트",      "sector": "유통"},
    {"ticker": "004170", "name": "신세계",      "dart_name": "신세계",      "sector": "유통"},
    {"ticker": "023530", "name": "롯데쇼핑",    "dart_name": "롯데쇼핑",    "sector": "유통"},
    {"ticker": "008770", "name": "호텔신라",    "dart_name": "호텔신라",    "sector": "유통"},
    {"ticker": "069960", "name": "현대백화점",  "dart_name": "현대백화점",  "sector": "유통"},
    {"ticker": "057050", "name": "현대홈쇼핑",  "dart_name": "현대홈쇼핑",  "sector": "유통"},
    {"ticker": "005930", "name": "삼성전자",    "dart_name": "삼성전자",    "sector": "반도체"},
    {"ticker": "000660", "name": "SK하이닉스",  "dart_name": "SK하이닉스", "sector": "반도체"},
    {"ticker": "042700", "name": "한미반도체",  "dart_name": "한미반도체",  "sector": "반도체장비"},
    {"ticker": "035420", "name": "NAVER",       "dart_name": "NAVER",       "sector": "IT플랫폼"},
    {"ticker": "035720", "name": "카카오",      "dart_name": "카카오",      "sector": "IT플랫폼"},
    {"ticker": "006400", "name": "삼성SDI",     "dart_name": "삼성SDI",     "sector": "배터리"},
]

# ── DART API 키 없을 때 사용할 정적 기준값 (공개 연간보고서 기준 추정) ───────
STATIC_TECH = {
    "139480": {"rd_billion": 23,      "rd_pct": 0.10, "patent_count": 11,    "tech_tags": ["물류자동화", "신선식품IT", "유통플랫폼"]},
    "004170": {"rd_billion": 15,      "rd_pct": 0.08, "patent_count": 8,     "tech_tags": ["패션테크", "면세IT", "리테일AI"]},
    "023530": {"rd_billion": 45,      "rd_pct": 0.15, "patent_count": 19,    "tech_tags": ["리테일테크", "물류자동화", "이커머스"]},
    "008770": {"rd_billion": 5,       "rd_pct": 0.05, "patent_count": 3,     "tech_tags": ["면세IT", "호텔AI", "서비스플랫폼"]},
    "069960": {"rd_billion": 12,      "rd_pct": 0.07, "patent_count": 6,     "tech_tags": ["리테일AI", "VR쇼핑", "물류"]},
    "057050": {"rd_billion": 8,       "rd_pct": 0.12, "patent_count": 4,     "tech_tags": ["이커머스", "라이브커머스", "AI추천"]},
    "005930": {"rd_billion": 280000,  "rd_pct": 8.20, "patent_count": 67000, "tech_tags": ["GAA공정", "HBM", "파운드리", "갤럭시AI", "엑시노스"]},
    "000660": {"rd_billion": 58000,   "rd_pct": 9.10, "patent_count": 21000, "tech_tags": ["HBM3E", "DRAM", "NAND", "AI메모리", "CXL"]},
    "042700": {"rd_billion": 1200,    "rd_pct": 5.30, "patent_count": 340,   "tech_tags": ["TC본더", "FC본더", "비전시스템", "반도체장비"]},
    "035420": {"rd_billion": 21000,   "rd_pct": 18.5, "patent_count": 3800,  "tech_tags": ["하이퍼클로바X", "생성AI", "검색AI", "클라우드", "로보틱스"]},
    "035720": {"rd_billion": 12000,   "rd_pct": 14.2, "patent_count": 1900,  "tech_tags": ["카카오브레인", "AI챗봇", "금융AI", "엔터테크"]},
    "006400": {"rd_billion": 65000,   "rd_pct": 8.30, "patent_count": 4700,  "tech_tags": ["전고체배터리", "하이니켈NCA", "46파이원통형", "실리콘음극재", "ESS"]},
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

        static = STATIC_TECH.get(ticker, {})
        entry  = {
            "ticker":      ticker,
            "name":        name,
            "sector":      stk["sector"],
            "rd_billion":  static.get("rd_billion"),   # 억원 단위
            "rd_pct":      static.get("rd_pct"),        # 매출 대비 %
            "patent_count": static.get("patent_count"),
            "tech_tags":   static.get("tech_tags", []),
            "disclosures": [],
            "data_quality": "static",
        }

        # DART API 수집 시도
        if DART_API_KEY and ticker in corp_map:
            corp_code = corp_map[ticker]
            print(f"    corp_code={corp_code}")

            rd, rev = get_rd_expense(corp_code, year=2024)
            if rd is not None:
                entry["rd_billion"]    = round(rd / 100)     # 백만원 → 억원
                entry["rd_pct"]        = round(rd / rev * 100, 2) if rev else None
                entry["data_quality"]  = "dart_api"
                print(f"    R&D: {entry['rd_billion']:,}억원 ({entry['rd_pct']}%)")
            else:
                print(f"    R&D: DART 응답 없음 — 정적값 사용")

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
