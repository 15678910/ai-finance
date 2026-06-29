"""
관심종목 모니터 — 대통령 'AI·반도체 메가프로젝트' 발표(2026-06-29 14:00) 수혜 테마 바스켓
==================================================================================
이미지의 3대 테마 × 세부 분야별 종목을 실시간 시세로 추적.
종목코드는 네이버 검색으로 1:1 검증해 하드코딩(런타임 모호성 차단 — 예: 'ISC'는
KISCO홀딩스가 아니라 반도체 테스트소켓 095340). 시세만 네이버 실시간 폴링 API로 조회.

출력: docs/watchlist.json
🚨 정보 모니터링용. 정책 테마주는 변동성 큼 — 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "watchlist.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

# 대통령 발표 예측 수혜 테마 — (대분류, 이모지, [ (소분류, [ (종목명, 코드 or None=비상장) ]) ])
# 코드는 네이버 검색으로 종목명 1:1 검증 완료(2026-06-29). None=비상장(시세 미표기).
THEMES = [
    ("반도체 메가프로젝트", "🔧", [
        ("종합 반도체", [("삼성전자", "005930"), ("SK하이닉스", "000660")]),
        ("소재·부품·장비", [("한미반도체", "042700"), ("HPSP", "403870"),
                       ("ISC", "095340"), ("리노공업", "058470"), ("원익IPS", "240810")]),
        ("유리기판·첨단 패키징", [("필옵틱스", "161580"), ("와이씨켐", "112290"), ("SKC", "011790")]),
    ]),
    ("피지컬 AI & 로봇", "🤖", [
        ("산업용·협동 로봇", [("레인보우로보틱스", "277810"), ("두산로보틱스", "454910"), ("로보티즈", "108490")]),
        ("AI 소프트웨어·엔진", [("솔트룩스", "304100"), ("코난테크놀로지", "402030"), ("마음AI", "377480")]),
    ]),
    ("기가와트(GW)급 AI 데이터센터", "⚡", [
        ("전력 인프라·변압기", [("HD현대일렉트릭", "267260"), ("효성중공업", "298040"), ("산일전기", "062040")]),
        ("전선·송배전", [("LS ELECTRIC", "010120"), ("LS전선", None), ("대한전선", "001440")]),
        ("냉각·데이터센터 설비", [("GST", "083450"), ("케이엔솔", "053080"), ("신성이엔지", "011930")]),
    ]),
]


def fetch_quotes(codes):
    """네이버 실시간 폴링 API 배치 조회 → {code: dict}."""
    url = "https://polling.finance.naver.com/api/realtime/domestic/stock/" + ",".join(codes)
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Referer": "https://finance.naver.com/"})
    raw = urllib.request.urlopen(req, timeout=15).read().decode("utf-8", errors="replace")
    out = {}
    for d in json.loads(raw).get("datas", []):
        out[d["itemCode"]] = d
    return out


def _num(d, key):
    """...Raw 필드 우선(숫자), 없으면 콤마 포함 문자열에서 파싱."""
    v = d.get(key + "Raw")
    if v is not None:
        try:
            return float(v)
        except Exception:
            pass
    s = str(d.get(key, "")).replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    all_codes = [c for _, _, subs in THEMES for _, lst in subs for _, c in lst if c]
    try:
        q = fetch_quotes(all_codes)
    except Exception as e:
        print(f"[ERROR] 시세 조회 실패: {e}")
        return 1

    market_status = None
    out_themes = []
    for theme, emoji, subs in THEMES:
        out_subs = []
        theme_chgs = []
        for sub, lst in subs:
            stocks = []
            for name, code in lst:
                if not code:
                    stocks.append({"name": name, "code": None, "listed": False,
                                   "note": "비상장(모회사 LS·관계사 LS전선아시아 229640)"})
                    continue
                d = q.get(code)
                if not d:
                    stocks.append({"name": name, "code": code, "listed": True, "price": None})
                    continue
                chg = _num(d, "fluctuationsRatio")
                if chg is not None:
                    theme_chgs.append(chg)
                market_status = market_status or d.get("marketStatus")
                stocks.append({
                    "name": d.get("stockName") or name, "code": code, "listed": True,
                    "price": _num(d, "closePrice"),
                    "change_pct": chg,
                    "change_amt": _num(d, "compareToPreviousClosePrice"),
                    "open": _num(d, "openPrice"), "high": _num(d, "highPrice"), "low": _num(d, "lowPrice"),
                    "volume": _num(d, "accumulatedTradingVolume"),
                    "mcap": d.get("marketValueFull"),   # 시가총액(억 단위 문자열)
                    "status": d.get("marketStatus"),
                })
            out_subs.append({"sub": sub, "stocks": stocks})
        avg = round(sum(theme_chgs) / len(theme_chgs), 2) if theme_chgs else None
        out_themes.append({"theme": theme, "emoji": emoji, "avg_change_pct": avg,
                           "n": len(theme_chgs), "subs": out_subs})
        print(f"{emoji} {theme}: 평균 {avg}% ({len(theme_chgs)}종목)")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "market_status": market_status,
        "event": {
            "title": "대통령 '대도약 3대 메가프로젝트' 발표 (반도체·피지컬AI·데이터센터)",
            "when": "2026-06-29 14:00 KST · 발표 완료",
            "scale": "향후 10년 최대 2,000조원 투자",
            "regions": "반도체=호남(제2 클러스터, 삼성·SK 참여) · AI데이터센터=충청·강원 · 피지컬AI벨트=영남",
            "support": "전력·용수·부지 공급 + 인허가 단축 + 전기요금 체계 개편 + 청와대 직할 담당관",
            "themes": ["반도체 메가프로젝트", "피지컬 AI·로봇", "GW급 AI 데이터센터"],
        },
        "themes": out_themes,
        "note": ("발표 예측 수혜 테마 바스켓(이미지 출처: 사용자 제공 브리핑). 종목코드 네이버 검증·하드코딩, "
                 "시세는 네이버 실시간 폴링. 정책 테마주는 발표 전후 변동성 극심 — 정보용·투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}  (시장상태 {market_status})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
