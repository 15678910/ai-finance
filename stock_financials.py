"""
종목 재무 스냅샷 — PER·PBR·ROE·배당 + 1년 밸류에이션 밴드 (Valley '종목 재무분석' 대응)
=======================================================================================
pykrx(KRX 로그인)로 종목별 기초 재무지표를 수집해 추세추종 보드의 '재무' 뷰에 공급.
  · PER/PBR: 현재값 + 최근 1년 밴드(최소~최대)와 백분위 → '역사적으로 싼가/비싼가'
  · ROE 근사 = EPS/BPS×100 (재무제표 기반 아님 — 근사치)
  · 배당수익률·DPS·EPS·BPS·시가총액·52주 위치
※ KRX 로그인 필요 → 내 PC(한국 IP) 로컬 스케줄러 전용.

출력: docs/stock_financials.json
🚨 정보 제공용 · 투자자문 아님.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "stock_financials.json")

WATCH = [
    ("000660", "SK하이닉스"),
    ("108490", "로보티즈"),
    ("005930", "삼성전자"),
    ("003550", "LG"),
    ("066570", "LG전자"),
    ("042700", "한미반도체"),
    ("009150", "삼성전기"),
    ("373220", "LG에너지솔루션"),
    ("247540", "에코프로비엠"),
]


def band(series):
    """(현재, 1년 최소, 최대, 백분위) — 0/NaN 제외."""
    vals = [float(v) for v in series if v and float(v) > 0]
    if not vals:
        return None
    cur = vals[-1]
    lo, hi = min(vals), max(vals)
    below = sum(1 for v in vals if v < cur)
    return {"cur": round(cur, 2), "lo": round(lo, 2), "hi": round(hi, 2),
            "pctile": round(below / len(vals) * 100)}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    if not (os.environ.get("KRX_ID") and os.environ.get("KRX_PW")):
        print("[INFO] KRX_ID/KRX_PW 미설정 — 수집 불가(기존 파일 보존).")
        return 0
    try:
        from pykrx import stock
        from pykrx.website.comm.auth import login_krx
        if not login_krx(os.environ["KRX_ID"], os.environ["KRX_PW"]):
            print("[ERROR] KRX 로그인 거부 — 중단.")
            return 1
    except ImportError:
        pass
    except Exception as e:
        print(f"[ERROR] pykrx 준비 실패: {e}")
        return 1
    from pykrx import stock

    now = datetime.now(KST)
    frm = (now - timedelta(days=370)).strftime("%Y%m%d")
    to = now.strftime("%Y%m%d")

    stocks, asof = [], None
    for code, name in WATCH:
        try:
            f = stock.get_market_fundamental_by_date(frm, to, code)   # BPS PER PBR EPS DIV DPS
            ohl = stock.get_market_ohlcv_by_date(frm, to, code)
            cap = stock.get_market_cap_by_date((now - timedelta(days=10)).strftime("%Y%m%d"), to, code)
        except Exception as e:
            print(f"  [WARN] {name}({code}) 조회 실패: {e}")
            continue
        if f is None or len(f) == 0 or ohl is None or len(ohl) == 0:
            print(f"  [WARN] {name}({code}) 데이터 없음")
            continue

        last = ohl.index[-1]
        asof = last.strftime("%Y-%m-%d") if hasattr(last, "strftime") else str(last)[:10]
        price = float(ohl["종가"].iloc[-1])
        hi52 = float(ohl["고가"].max())
        lo52 = float(ohl["저가"].min())
        w52_pos = round((price - lo52) / (hi52 - lo52) * 100) if hi52 > lo52 else None

        per_b = band(f["PER"].tolist())
        pbr_b = band(f["PBR"].tolist())
        eps = float(f["EPS"].iloc[-1]) if "EPS" in f.columns else None
        bps = float(f["BPS"].iloc[-1]) if "BPS" in f.columns else None
        div = float(f["DIV"].iloc[-1]) if "DIV" in f.columns else None
        dps = float(f["DPS"].iloc[-1]) if "DPS" in f.columns else None
        roe = round(eps / bps * 100, 1) if (eps and bps and bps > 0) else None
        mcap = None
        try:
            mcap = int(cap["시가총액"].iloc[-1])
        except Exception:
            pass

        stocks.append({
            "code": code, "name": name, "asof": asof, "price": round(price),
            "per": per_b, "pbr": pbr_b,
            "eps": round(eps) if eps else None, "bps": round(bps) if bps else None,
            "roe_approx": roe, "div_yield": div, "dps": round(dps) if dps else None,
            "mcap": mcap, "w52": {"hi": round(hi52), "lo": round(lo52), "pos": w52_pos},
        })
        print(f"  {name}: PER {per_b['cur'] if per_b else '—'} (1y {per_b['lo']}~{per_b['hi']}, 백분위 {per_b['pctile']}) "
              f"PBR {pbr_b['cur'] if pbr_b else '—'} ROE≈{roe}% 배당 {div}% 52주 {w52_pos}%" if per_b else f"  {name}: PER 없음(적자 등)")

    if not stocks:
        print("[ERROR] 전 종목 실패 — 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": asof, "stocks": stocks,
        "note": ("KRX 종목 기초지표(일별): PER/PBR 1년 밴드·백분위(낮을수록 역사적 저평가), "
                 "ROE≈EPS/BPS 근사(재무제표 기반 아님), DIV=배당수익률(%), 시총=원. "
                 "적자 기업은 PER 미표시. 정보 제공용 · 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f2:
        json.dump(out, f2, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE} ({len(stocks)}종목 · 기준 {asof})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
