"""
퀀트 스크리너 — 가치 + 성장 + 고배당 (GARP·배당 조건 필터)
============================================================
KRX 전 종목(코스피+코스닥)에서 아래 조건을 '기계적으로' 걸러 나열한다.
  · 가치: 0 < PER ≤ 12  그리고  0 < PBR ≤ 1.5
  · 성장: EPS 전년 대비 +10% 이상 (양수→양수)  그리고  ROE(≈EPS/BPS) ≥ 8%
  · 배당: 배당수익률(DIV) ≥ 3%
  · 규모: 시가총액 ≥ 3,000억 (소형주 노이즈 제거)
정렬: 배당률 내림차순 · 상위 20종목.
※ KRX 로그인 필요(로컬 스케줄러 전용). EPS/배당은 후행 지표 — 미래 보장 없음.

출력: docs/stock_screener.json
🚨 조건 충족 종목의 '나열'이며 종목 추천이 아님 · 투자자문 아님 · 밸류트랩 유의.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "stock_screener.json")

CRIT = {"per_max": 12.0, "pbr_max": 1.5, "epsg_min": 10.0, "roe_min": 8.0,
        "div_min": 3.0, "cap_min": 3000e8}
TOP_N = 20


def pick_trading_day(stock, base):
    """base 근처의 유효 거래일(펀더멘털 EPS 유효 종목 100+) 탐색."""
    for k in range(0, 10):
        d = (base - timedelta(days=k)).strftime("%Y%m%d")
        try:
            f = stock.get_market_fundamental_by_ticker(d, market="KOSPI")
            if f is not None and len(f) > 100 and (f["EPS"] != 0).sum() > 100:
                return d, None
        except Exception as e:
            last = e
            continue
    return None, "유효 거래일 탐색 실패"


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
        from pykrx import stock

    now = datetime.now(KST)
    d_now, err = pick_trading_day(stock, now)
    d_prev, err2 = pick_trading_day(stock, now - timedelta(days=365))
    if not d_now or not d_prev:
        print(f"[ERROR] 거래일 탐색 실패: {err or err2}")
        return 1
    print(f"  기준일 {d_now} · 1년 전 {d_prev}")

    rows = []
    funnel = {"universe": 0, "value": 0, "growth": 0, "div": 0, "cap": 0}
    for mkt in ("KOSPI", "KOSDAQ"):
        try:
            f = stock.get_market_fundamental_by_ticker(d_now, market=mkt)
            fp = stock.get_market_fundamental_by_ticker(d_prev, market=mkt)
            cap = stock.get_market_cap_by_ticker(d_now, market=mkt)
        except Exception as e:
            print(f"  [WARN] {mkt} 조회 실패: {e}")
            continue
        funnel["universe"] += len(f)
        for code, r in f.iterrows():
            per, pbr = float(r["PER"]), float(r["PBR"])
            eps, bps = float(r["EPS"]), float(r["BPS"])
            div, dps = float(r["DIV"]), float(r["DPS"])
            # 가치
            if not (0 < per <= CRIT["per_max"] and 0 < pbr <= CRIT["pbr_max"]):
                continue
            funnel["value"] += 1
            # 성장 (EPS 양수→양수 & +10%↑, ROE≥8%)
            eps_prev = float(fp.loc[code, "EPS"]) if code in fp.index else 0.0
            if eps_prev <= 0 or eps <= 0:
                continue
            epsg = (eps / eps_prev - 1) * 100
            roe = eps / bps * 100 if bps > 0 else 0
            if epsg < CRIT["epsg_min"] or roe < CRIT["roe_min"]:
                continue
            funnel["growth"] += 1
            # 배당
            if div < CRIT["div_min"]:
                continue
            funnel["div"] += 1
            # 규모
            c = float(cap.loc[code, "시가총액"]) if code in cap.index else 0
            price = float(cap.loc[code, "종가"]) if code in cap.index else None
            if c < CRIT["cap_min"]:
                continue
            funnel["cap"] += 1
            rows.append({"code": code, "market": mkt, "price": round(price) if price else None,
                         "cap_jo": round(c / 1e12, 2),
                         "per": round(per, 1), "pbr": round(pbr, 2),
                         "eps_growth": round(epsg), "roe": round(roe, 1),
                         "div": round(div, 2), "dps": round(dps)})

    rows.sort(key=lambda x: -x["div"])
    rows = rows[:TOP_N]
    for r in rows:                                   # 이름은 통과 종목만 조회
        try:
            r["name"] = stock.get_market_ticker_name(r["code"])
        except Exception:
            r["name"] = r["code"]
        print(f"  {r['name']}({r['code']}) {r['market']} | PER {r['per']} PBR {r['pbr']} "
              f"ROE {r['roe']}% EPS성장 +{r['eps_growth']}% 배당 {r['div']}% 시총 {r['cap_jo']}조")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": f"{d_now[:4]}-{d_now[4:6]}-{d_now[6:]}",
        "prev_date": f"{d_prev[:4]}-{d_prev[4:6]}-{d_prev[6:]}",
        "criteria": {"PER": f"≤{CRIT['per_max']:.0f}", "PBR": f"≤{CRIT['pbr_max']}",
                     "EPS성장(YoY)": f"≥{CRIT['epsg_min']:.0f}%", "ROE(근사)": f"≥{CRIT['roe_min']:.0f}%",
                     "배당수익률": f"≥{CRIT['div_min']:.0f}%", "시가총액": "≥3,000억"},
        "funnel": funnel,
        "stocks": rows,
        "note": ("코스피+코스닥 전 종목을 공개 기준(가치 PER·PBR / 성장 EPS YoY·ROE근사 / 배당 DIV / 규모)으로 "
                 "기계적으로 필터링한 '나열'이며 종목 추천이 아님. EPS·배당은 후행 지표(미래 미보장), "
                 "ROE=EPS/BPS 근사, 배당률 높은 종목은 배당컷·밸류트랩 위험 검토 필수. "
                 "KRX 일별 지표 · 투자 판단은 본인 책임."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as fo:
        json.dump(out, fo, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE} ({len(rows)}종목 · 퍼널 {funnel})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
