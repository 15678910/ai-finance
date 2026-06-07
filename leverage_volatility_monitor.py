"""
레버리지·변동성 모니터
========================
ETF·레버리지가 시장 변동성을 증폭하는 미시구조를 추적:

  ① 레버리지/인버스 ETF 추적 (AUM·등락·추정 마감 리밸런싱 압력)
  ② 레버리지 ÷ 인버스 자금 비율 (개인 포지셔닝 쏠림 = 역발상 심리)
  ③ 변동성 레짐 신호 (VIX 수준 + 추세 → 정상/경계/위험 + 캐스케이드 경고)
  ④ 국면별 대응 가이드

출력: docs/leverage_volatility.json

🚨 시뮬레이션·분석용. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "docs", "data.json")
OVERSEAS_FILE = os.path.join(BASE_DIR, "docs", "overseas_market.json")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "leverage_volatility.json")

# 주요 한국 레버리지·인버스 ETF (KRX)
LEVERAGE_ETFS = [
    {"ticker": "122630", "name": "KODEX 레버리지",        "type": "leverage", "mult": 2.0, "underlying": "KOSPI200"},
    {"ticker": "252670", "name": "KODEX 200선물인버스2X", "type": "inverse",  "mult": -2.0, "underlying": "KOSPI200"},
    {"ticker": "114800", "name": "KODEX 인버스",          "type": "inverse",  "mult": -1.0, "underlying": "KOSPI200"},
    {"ticker": "233740", "name": "KODEX 코스닥150레버리지","type": "leverage", "mult": 2.0, "underlying": "KOSDAQ150"},
    {"ticker": "251340", "name": "KODEX 코스닥150선물인버스","type": "inverse","mult": -1.0, "underlying": "KOSDAQ150"},
    {"ticker": "123320", "name": "TIGER 레버리지",        "type": "leverage", "mult": 2.0, "underlying": "KOSPI200"},
]

# 정적 폴백 AUM (조원, 2026-06 기준 추정)
STATIC_AUM = {
    "122630": 2.85, "252670": 2.10, "114800": 0.62,
    "233740": 0.95, "251340": 0.48, "123320": 0.18,
}

# 해외(미국) 레버리지·인버스 ETF (USD, AUM 단위 $B)
OVERSEAS_LEVERAGE_ETFS = [
    {"ticker": "TQQQ", "name": "ProShares 3X 나스닥100",   "type": "leverage", "mult": 3.0,  "underlying": "나스닥100"},
    {"ticker": "SQQQ", "name": "ProShares -3X 나스닥100",  "type": "inverse",  "mult": -3.0, "underlying": "나스닥100"},
    {"ticker": "SOXL", "name": "Direxion 3X 반도체",       "type": "leverage", "mult": 3.0,  "underlying": "반도체(SOX)"},
    {"ticker": "SOXS", "name": "Direxion -3X 반도체",      "type": "inverse",  "mult": -3.0, "underlying": "반도체(SOX)"},
    {"ticker": "NVDL", "name": "GraniteShares 2X 엔비디아","type": "leverage", "mult": 2.0,  "underlying": "엔비디아"},
    {"ticker": "TSLL", "name": "Direxion 2X 테슬라",       "type": "leverage", "mult": 2.0,  "underlying": "테슬라"},
    {"ticker": "UPRO", "name": "ProShares 3X S&P500",      "type": "leverage", "mult": 3.0,  "underlying": "S&P500"},
    {"ticker": "SPXU", "name": "ProShares -3X S&P500",     "type": "inverse",  "mult": -3.0, "underlying": "S&P500"},
]

# 정적 폴백 AUM ($B, 2026-06 기준 추정)
STATIC_AUM_USD = {
    "TQQQ": 26.0, "SQQQ": 3.5, "SOXL": 13.0, "SOXS": 1.2,
    "NVDL": 6.0, "TSLL": 4.0, "UPRO": 4.0, "SPXU": 0.6,
}


def fetch_overseas_etf(ticker: str, name: str) -> dict:
    """yfinance로 해외 ETF 가격·등락 수집 (AUM은 정적 폴백)."""
    out = {"price": None, "change_pct": 0.0, "aum_b": STATIC_AUM_USD.get(ticker)}
    try:
        import yfinance as yf
        t = yf.Ticker(ticker)
        fi = t.fast_info
        price = getattr(fi, "last_price", None)
        prev = getattr(fi, "previous_close", None)
        if price:
            out["price"] = round(float(price), 2)
            if prev and float(prev) > 0:
                out["change_pct"] = round((float(price) - float(prev)) / float(prev) * 100, 2)
    except Exception as e:
        print(f"  [WARN] {name} 가격 수집 실패: {e}")
    return out


def fetch_etf(ticker: str, name: str, today_str: str) -> dict:
    """yfinance + pykrx로 ETF 가격·등락·AUM 수집."""
    out = {"price": None, "change_pct": 0.0, "aum_tril": STATIC_AUM.get(ticker)}
    # 가격·등락 (yfinance)
    try:
        import yfinance as yf
        for suffix in [".KS", ".KQ"]:
            t = yf.Ticker(ticker + suffix)
            fi = t.fast_info
            price = getattr(fi, "last_price", None)
            prev = getattr(fi, "previous_close", None)
            if price:
                out["price"] = round(float(price), 0)
                if prev and float(prev) > 0:
                    out["change_pct"] = round((float(price) - float(prev)) / float(prev) * 100, 2)
                break
    except Exception as e:
        print(f"  [WARN] {name} 가격 수집 실패: {e}")
    # AUM은 별도 API 필요 → 정적값 유지 (today_str 예약)
    _ = today_str
    return out


def get_vix() -> dict:
    """VIX 수준 + 추세 (overseas_market.json 또는 data.json)."""
    vix_cur, vix_chg = None, None
    # overseas
    try:
        with open(OVERSEAS_FILE, encoding="utf-8") as f:
            om = json.load(f)
        for m in om.get("all_markets", []):
            if m.get("ticker") == "^VIX" or m.get("is_vix"):
                vix_cur = m.get("current")
                vix_chg = m.get("change_pct")
                break
    except Exception:
        pass
    # data.json macro_detail 폴백
    if vix_cur is None:
        try:
            with open(DATA_FILE, encoding="utf-8") as f:
                d = json.load(f)
            v = d.get("macro_detail", {}).get("vix", {})
            if isinstance(v, dict):
                vix_cur = v.get("current")
        except Exception:
            pass
    return {"current": vix_cur, "change_pct": vix_chg}


def classify_vol_regime(vix) -> dict:
    """VIX 수준 → 변동성 레짐 + 대응 가이드."""
    try:
        vix = float(vix) if vix is not None else None
    except (ValueError, TypeError):
        vix = None
    if vix is None:
        return {"level": "데이터없음", "color": "muted", "cascade_risk": "—",
                "guide": ["VIX 데이터 수집 불가"]}
    if vix < 15:
        return {"level": "안정", "color": "green", "cascade_risk": "낮음",
                "guide": ["정상 국면. 레버리지 ETF 단기 활용 가능하나 음의 복리 유의",
                          "장 마감 리밸런싱 변동은 제한적"]}
    if vix < 20:
        return {"level": "정상", "color": "cyan", "cascade_risk": "낮음",
                "guide": ["평상 변동성. 레버리지 ETF 장기 보유는 여전히 비권장",
                          "포지션 정상 유지"]}
    if vix < 25:
        return {"level": "경계", "color": "amber", "cascade_risk": "중간",
                "guide": ["변동성 상승 국면. 변동성 타겟팅 펀드 디레버리징 시작 가능",
                          "레버리지 ETF 음의 복리 가속 — 보유 축소 검토",
                          "장 마감 30분 기계적 변동 확대 주의"]}
    if vix < 30:
        return {"level": "위험", "color": "red", "cascade_risk": "높음",
                "guide": ["고변동성. 매도가 매도를 부르는 디레버리징 캐스케이드 위험",
                          "레버리지·곱버스 즉시 축소. 포지션 사이즈 절반 이하 권고",
                          "단순·방어적 전략으로 전환 (퀀트 보고서: 단순함의 힘)"]}
    return {"level": "패닉", "color": "red", "cascade_risk": "극심",
            "guide": ["시장 공포. 강제 청산·마진콜 연쇄 가능",
                      "레버리지 전량 청산. 현금 비중 확대",
                      "곱버스 쏠림 극단 시 역발상 바닥 신호 가능성 — 단, 무리한 진입 금지"]}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    today_str = datetime.now(KST).strftime("%Y%m%d")
    print("=" * 55)
    print("  레버리지·변동성 모니터")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    # ① ETF 추적
    print("\n[①] 레버리지·인버스 ETF 수집...")
    etfs = []
    lev_aum, inv_aum = 0.0, 0.0
    for e in LEVERAGE_ETFS:
        d = fetch_etf(e["ticker"], e["name"], today_str)
        entry = {**e, **d}
        # 추정 마감 리밸런싱 압력 = AUM × |배율-1| × 당일등락 (방향성 증폭분)
        aum = d.get("aum_tril") or 0
        chg = d.get("change_pct") or 0
        # 레버리지 ETF는 매일 (배율-1)배만큼 추가 익스포저 조정 필요
        rebal = round(aum * abs(e["mult"] - (1 if e["mult"] > 0 else -1)) * abs(chg) / 100, 4)
        entry["rebal_pressure_tril"] = rebal
        entry["rebal_direction"] = ("매수" if (chg > 0 and e["mult"] > 0) or (chg < 0 and e["mult"] < 0)
                                    else "매도") if chg != 0 else "—"
        etfs.append(entry)
        if e["type"] == "leverage":
            lev_aum += aum
        else:
            inv_aum += aum
        print(f"  {e['name']}: {d.get('price')} ({chg:+.2f}%) AUM {aum}조 리밸런싱 {rebal}조")

    # ①-b 해외 레버리지·인버스 ETF
    print("\n[①-b] 해외 레버리지·인버스 ETF 수집...")
    overseas_etfs = []
    o_lev, o_inv = 0.0, 0.0
    for e in OVERSEAS_LEVERAGE_ETFS:
        d = fetch_overseas_etf(e["ticker"], e["name"])
        entry = {**e, **d}
        aum = d.get("aum_b") or 0
        chg = d.get("change_pct") or 0
        rebal = round(aum * abs(e["mult"] - (1 if e["mult"] > 0 else -1)) * abs(chg) / 100, 3)
        entry["rebal_pressure_b"] = rebal
        entry["rebal_direction"] = ("매수" if (chg > 0 and e["mult"] > 0) or (chg < 0 and e["mult"] < 0)
                                    else "매도") if chg != 0 else "—"
        overseas_etfs.append(entry)
        if e["type"] == "leverage":
            o_lev += aum
        else:
            o_inv += aum
        print(f"  {e['name']}: {d.get('price')} ({chg:+.2f}%) AUM ${aum}B 리밸런싱 ${rebal}B")
    o_ratio = round(o_lev / o_inv, 2) if o_inv > 0 else None
    if o_ratio is not None:
        if o_ratio > 4:
            o_signal = "레버리지(롱) 극단 쏠림 — 미국 과열·고점 경계 (역발상: 하락 대비)"
        elif o_ratio < 1.5:
            o_signal = "인버스 비중 확대 — 헤지·비관 심리 (역발상: 바닥 신호 가능)"
        else:
            o_signal = "레버리지 우위 — 정상 범위 (강세장 통상)"
    else:
        o_signal = "데이터 부족"
    print(f"  해외 레버리지/인버스 비율: {o_ratio} → {o_signal}")

    # ② 레버리지/인버스 비율
    ratio = round(lev_aum / inv_aum, 2) if inv_aum > 0 else None
    if ratio is not None:
        if ratio > 1.5:
            ratio_signal = "레버리지(롱) 쏠림 — 시장 과열·고점 경계 (역발상: 하락 대비)"
        elif ratio < 0.7:
            ratio_signal = "인버스(곱버스) 쏠림 — 비관 극단 (역발상: 바닥 신호 가능)"
        else:
            ratio_signal = "중립 — 포지셔닝 균형"
    else:
        ratio_signal = "데이터 부족"
    print(f"\n[②] 레버리지/인버스 비율: {ratio} → {ratio_signal}")

    # ③ 변동성 레짐
    vix = get_vix()
    regime = classify_vol_regime(vix.get("current"))
    print(f"\n[③] VIX {vix.get('current')} → {regime['level']} (캐스케이드 위험: {regime['cascade_risk']})")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "etfs": etfs,
        "leverage_aum_tril": round(lev_aum, 2),
        "inverse_aum_tril": round(inv_aum, 2),
        "lev_inv_ratio": ratio,
        "ratio_signal": ratio_signal,
        "overseas_etfs": overseas_etfs,
        "overseas_leverage_aum_b": round(o_lev, 1),
        "overseas_inverse_aum_b": round(o_inv, 1),
        "overseas_ratio": o_ratio,
        "overseas_ratio_signal": o_signal,
        "vix": vix,
        "vol_regime": regime,
        "mechanisms": [
            {"name": "레버리지 ETF 일일 리밸런싱", "desc": "2X·3X ETF가 매일 배율 유지 위해 상승 시 추가매수·하락 시 추가매도 → 추세 증폭 (특히 장 마감 30분)"},
            {"name": "음의 복리(변동성 붕괴)", "desc": "등락 반복 시 레버리지 ETF가 지수보다 더 깎임 → 장기 보유 손실·청산 매물"},
            {"name": "변동성 타겟팅 디레버리징", "desc": "변동성 상승 시 리스크패리티·vol-control 펀드 자동 매도 → 매도가 매도를 부르는 피드백"},
            {"name": "패시브 ETF 강제 매매", "desc": "지수 편출입·자금 유출입에 펀더멘털 무관 기계적 주문 (SpaceX IPO 사례)"},
        ],
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
