"""
AI 패러다임 모니터 (성상현 부부장 브리핑 기반)
================================================
① MV=PY 생산성 진단 — M2·유통속도(V)·물가(P)·실질성장(Y)으로
   디스인플레이션 성장 사이클 vs 인플레 우위를 판별
② 휴머노이드·Physical AI 테마 트래커 (LG전자·SK하이닉스 포함)
③ 하이퍼스케일러 CAPEX 버블 게이지 (닷컴 35% 대비)

출력: docs/ai_paradigm.json
🚨 시뮬레이션·분석용. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
import urllib.request
import urllib.parse
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "ai_paradigm.json")
USER_AGENT = "Mozilla/5.0 (compatible; ai-finance-paradigm/1.0)"


# ── FRED 원시값 수집 (units= 미사용, 400 회피) ────────────────────────
def fred_observations(api_key: str, series_id: str, years_back: int = 3) -> list:
    """FRED 원시 관측값 [(date, value), ...] 오름차순."""
    if not api_key:
        return []
    try:
        start = f"{date.today().year - years_back}-01-01"
        params = urllib.parse.urlencode({
            "series_id": series_id, "api_key": api_key.strip(),
            "observation_start": start, "file_type": "json", "sort_order": "asc",
        })
        url = f"https://api.stlouisfed.org/fred/series/observations?{params}"
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=15) as r:
            d = json.loads(r.read())
        return [(o["date"], float(o["value"])) for o in d.get("observations", [])
                if o.get("value") not in (".", "", None)]
    except Exception as e:
        print(f"  [WARN] FRED {series_id} 실패: {e}")
        return []


def _yoy_latest(rows: list):
    """최신 전년동월비 % + 기준일."""
    if len(rows) < 13:
        return None, None
    d2v = {d: v for d, v in rows}
    d, v = rows[-1]
    ya = f"{int(d[:4]) - 1}{d[4:]}"
    pv = d2v.get(ya)
    if pv and pv != 0:
        return round((v - pv) / pv * 100, 2), d[:7]
    return None, None


# ① MV=PY 생산성 진단 ────────────────────────────────────────────────
def build_mv_py(api_key: str) -> dict:
    print("\n[①] MV=PY 생산성 진단...")
    # M: M2 YoY
    m2 = fred_observations(api_key, "M2SL")
    m_yoy, m_date = _yoy_latest(m2)
    # V: M2 유통속도 (FRED M2V, 분기) — 최신값 + 전년대비
    m2v = fred_observations(api_key, "M2V", years_back=3)
    v_now = m2v[-1][1] if m2v else None
    v_date = m2v[-1][0][:7] if m2v else None
    v_yoy = None
    if len(m2v) >= 5:  # 분기 데이터 4개 전 = 1년 전
        prev = m2v[-5][1]
        if prev:
            v_yoy = round((m2v[-1][1] - prev) / prev * 100, 2)
    # P: CPI YoY
    cpi = fred_observations(api_key, "CPIAUCSL")
    p_yoy, p_date = _yoy_latest(cpi)
    # Y: 실질 GDP 성장률 (연율, 분기) A191RL1Q225SBEA
    gdp = fred_observations(api_key, "A191RL1Q225SBEA", years_back=3)
    y_now = gdp[-1][1] if gdp else None
    y_date = gdp[-1][0][:7] if gdp else None

    # 판정
    verdict, vcolor, vdesc = "데이터 부족", "muted", "FRED 데이터 수집 필요"
    if p_yoy is not None and y_now is not None:
        if y_now > p_yoy and y_now > 0:
            verdict, vcolor = "디스인플레이션 성장", "green"
            vdesc = "실질성장(Y)이 물가(P)를 압도 — 브리핑의 '건전한 성장' 국면. 부채를 인플레 아닌 성장으로 해결."
        elif y_now > 0 and p_yoy > 0:
            verdict, vcolor = "성장-인플레 공존", "amber"
            vdesc = "성장과 물가가 함께 상승 — 생산성(Y) 가속이 P를 추월하는지 주시."
        elif y_now <= 0:
            verdict, vcolor = "성장 둔화", "red"
            vdesc = "실질성장 정체/위축 — 디스인플레이션 사이클 미도래."
    print(f"  M(M2): {m_yoy}% | V: {v_now} ({v_yoy}%) | P(CPI): {p_yoy}% | Y(GDP): {y_now}% → {verdict}")

    return {
        "M": {"label": "M2 통화량", "value": m_yoy, "unit": "% YoY", "date": m_date, "desc": "정부 재정·기업 투자로 팽창"},
        "V": {"label": "유통속도", "value": v_now, "yoy": v_yoy, "date": v_date, "desc": "AI 에이전트·로봇 자동결제로 가속 가능"},
        "P": {"label": "물가(CPI)", "value": p_yoy, "unit": "% YoY", "date": p_date, "desc": "통화량 증가 압력 vs 생산성 상쇄"},
        "Y": {"label": "실질성장(GDP)", "value": y_now, "unit": "% 연율", "date": y_date, "desc": "핵심변수 — AI·로봇이 폭발적 증가 견인"},
        "verdict": verdict, "verdict_color": vcolor, "verdict_desc": vdesc,
        "equation": "MV = PY (통화량×유통속도 = 물가×실질성장)",
    }


# ② 휴머노이드·Physical AI 테마 ───────────────────────────────────────
HUMANOID_STOCKS = [
    ("TSLA",      "테슬라",        "🇺🇸", "Optimus 휴머노이드 양산 추진"),
    ("NVDA",      "엔비디아",      "🇺🇸", "Isaac 로보틱스 플랫폼·AI 두뇌"),
    ("005380.KS", "현대차",        "🇰🇷", "보스턴다이내믹스·Atlas 휴머노이드"),
    ("066570.KS", "LG전자",        "🇰🇷", "CLOiD 가정용 로봇·Isaac 협력"),
    ("000660.KS", "SK하이닉스",   "🇰🇷", "HBM — 로봇 AI 연산 메모리 핵심"),
    ("005930.KS", "삼성전자",      "🇰🇷", "레인보우로보틱스 인수·HBM4E"),
    ("277810.KS", "레인보우로보틱스","🇰🇷", "삼성 휴머노이드 본체"),
    ("454910.KS", "두산로보틱스",  "🇰🇷", "협동로봇·산업 자동화"),
    ("056190.KS", "에스에프에이",  "🇰🇷", "스마트팩토리 자동화"),
    ("108490.KS", "로보티즈",      "🇰🇷", "로봇 구동 액추에이터"),
]


def fetch_stock(ticker: str) -> dict:
    try:
        import yfinance as yf
        out = {"price": None, "change_pct": 0.0, "market_cap": None}
        ylist = [ticker + ".KS", ticker + ".KQ"] if (ticker.isdigit()) else [ticker]
        if "." in ticker:
            ylist = [ticker]
        for yt in ylist:
            t = yf.Ticker(yt)
            fi = t.fast_info
            p = getattr(fi, "last_price", None)
            pv = getattr(fi, "previous_close", None)
            mc = getattr(fi, "market_cap", None)
            if p:
                out["price"] = round(float(p), 2)
                if pv and float(pv) > 0:
                    out["change_pct"] = round((float(p) - float(pv)) / float(pv) * 100, 2)
                if mc:
                    out["market_cap"] = round(float(mc) / 1e12, 3)
                return out
        return out
    except Exception as e:
        print(f"  [WARN] {ticker} 실패: {e}")
        return {"price": None, "change_pct": 0.0, "market_cap": None}


def build_humanoid() -> list:
    print("\n[②] 휴머노이드·Physical AI 테마...")
    rows = []
    for tk, name, flag, role in HUMANOID_STOCKS:
        d = fetch_stock(tk)
        rows.append({"ticker": tk.replace(".KS", "").replace(".KQ", ""),
                     "name": name, "flag": flag, "role": role, **d})
        print(f"  {name}: {d.get('price')} ({d.get('change_pct')}%)")
    return rows


# ③ 하이퍼스케일러 CAPEX 게이지 ───────────────────────────────────────
# yfinance capex 불안정 → 2026 추정 연간 CAPEX·매출 (정적, 단위 $B)
HYPERSCALER_CAPEX = [
    {"ticker": "MSFT",  "name": "마이크로소프트", "capex_b": 88,  "revenue_b": 290, "note": "Azure·OpenAI 인프라"},
    {"ticker": "GOOGL", "name": "알파벳",        "capex_b": 78,  "revenue_b": 385, "note": "TPU·데이터센터"},
    {"ticker": "AMZN",  "name": "아마존",        "capex_b": 105, "revenue_b": 660, "note": "AWS·Trainium"},
    {"ticker": "META",  "name": "메타",          "capex_b": 74,  "revenue_b": 185, "note": "Llama·AI 인프라"},
    {"ticker": "ORCL",  "name": "오라클",        "capex_b": 25,  "revenue_b": 60,  "note": "OCI 클라우드 급팽창"},
]


def build_capex() -> dict:
    print("\n[③] 하이퍼스케일러 CAPEX 게이지...")
    items = []
    total_capex, total_rev = 0, 0
    for h in HYPERSCALER_CAPEX:
        ratio = round(h["capex_b"] / h["revenue_b"] * 100, 1)
        items.append({**h, "capex_ratio": ratio})
        total_capex += h["capex_b"]
        total_rev += h["revenue_b"]
        print(f"  {h['name']}: CAPEX ${h['capex_b']}B / 매출 ${h['revenue_b']}B = {ratio}%")
    avg_ratio = round(total_capex / total_rev * 100, 1) if total_rev else 0
    items.sort(key=lambda x: x["capex_ratio"], reverse=True)
    return {
        "items": items,
        "total_capex_b": total_capex,
        "avg_ratio": avg_ratio,
        "dotcom_ref": 35,  # 브리핑: 닷컴 당시 IT CAPEX 비중 (인플레 조정)
        "signal": ("버블 경계 — 닷컴(35%) 수준 근접" if avg_ratio >= 25 else "정상 투자 사이클"),
    }


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    api_key = os.environ.get("FRED_API_KEY", "")
    print("=" * 55)
    print("  AI 패러다임 모니터 (MV=PY · 휴머노이드 · CAPEX)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  FRED_API_KEY: {'있음' if api_key else '없음'}")
    print("=" * 55)

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "source": "성상현 부부장 브리핑 — AI/휴머노이드 경제 패러다임",
        "mv_py": build_mv_py(api_key),
        "humanoid": build_humanoid(),
        "capex": build_capex(),
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
