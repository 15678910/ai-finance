"""
분봉 단타 보조지표 (SK하이닉스 · XRP)
=====================================
⚠️ 체결은 증권사 HTS/거래소로. 이건 '진입 전 분봉 맥락 참고'용 보조지표.
⚠️ yfinance 데이터는 15~20분 지연 가능 — 실시간 아님(데이터 시각 함께 표시).

지표(5분·15분봉):
  - VWAP 이격(%)   : 당일 거래량가중평균 대비 (단타 핵심 기준선)
  - RSI(7)        : 단기 과매수(70+)/과매도(30-)
  - z20           : 20봉 이동평균 대비 표준편차
  - 당일 위치(%)   : 당일 저가~고가 중 현재 위치

출력: docs/intraday.json
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "intraday.json")

TARGETS = [
    {"key": "SK하이닉스", "ticker": "000660.KS", "flag": "🇰🇷", "unit": "원"},
    {"key": "XRP",       "ticker": "XRP-USD",   "flag": "⚪", "unit": "$"},
]


def _rsi(series, n=7):
    d = series.diff()
    up = d.clip(lower=0).rolling(n).mean()
    dn = (-d.clip(upper=0)).rolling(n).mean()
    return 100 - 100 / (1 + up / dn)


def analyze_tf(df, np):
    """단일 타임프레임 DataFrame(OHLCV) → 지표 dict."""
    close = df["Close"].dropna()
    if len(close) < 25:
        return None
    # 당일 바만 (VWAP·당일위치)
    last_day = close.index[-1].normalize()
    today_mask = close.index.normalize() == last_day
    tclose = close[today_mask]
    thigh = df["High"][today_mask]
    tlow = df["Low"][today_mask]
    tvol = df["Volume"][today_mask].fillna(0)
    price = float(close.iloc[-1])

    # VWAP (당일 누적)
    if float(tvol.sum()) > 0:
        vwap = float((tclose * tvol).cumsum().iloc[-1] / tvol.cumsum().iloc[-1])
    else:
        vwap = float(tclose.mean())
    vwap_dev = (price / vwap - 1) * 100 if vwap else 0.0

    # RSI(7)
    rsi = _rsi(close, 7)
    rsi_now = float(rsi.iloc[-1]) if not np.isnan(rsi.iloc[-1]) else None

    # z20
    ma = close.rolling(20).mean()
    sd = close.rolling(20).std()
    z20 = float((close.iloc[-1] - ma.iloc[-1]) / sd.iloc[-1]) if sd.iloc[-1] else 0.0

    # 당일 위치
    hi, lo = float(thigh.max()), float(tlow.min())
    pos = (price - lo) / (hi - lo) * 100 if hi > lo else 50.0

    return {
        "price": round(price, 4 if price < 100 else 0),
        "vwap": round(vwap, 4 if vwap < 100 else 0),
        "vwap_dev": round(vwap_dev, 2),
        "rsi7": round(rsi_now, 0) if rsi_now is not None else None,
        "z20": round(z20, 2),
        "day_pos": round(pos, 0),
        "last_bar": close.index[-1].tz_convert(KST).strftime("%m-%d %H:%M") if close.index[-1].tzinfo else close.index[-1].strftime("%m-%d %H:%M"),
    }


def classify(tf5):
    """5분봉 기준 단기 신호."""
    if not tf5:
        return "데이터없음", "muted"
    rsi = tf5.get("rsi7"); z = tf5.get("z20"); vd = tf5.get("vwap_dev")
    if (rsi is not None and rsi <= 25) or (z is not None and z <= -2):
        return "🟢 단기 과매도 (반등 관심)", "green"
    if (rsi is not None and rsi >= 75) or (z is not None and z >= 2):
        return "🔴 단기 과매수 (차익 관심)", "red"
    if vd is not None and vd < -0.5:
        return "🔵 VWAP 하단 — 약세 흐름", "cyan"
    if vd is not None and vd > 0.5:
        return "🟠 VWAP 상단 — 강세 흐름", "amber"
    return "⚪ 중립 (VWAP 부근)", "muted"


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    import numpy as np
    import yfinance as yf
    import warnings
    warnings.filterwarnings("ignore")

    print("=" * 55)
    print("  분봉 단타 보조지표 (SK하이닉스 · XRP)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    def _flatten(d):
        if d is not None and hasattr(d.columns, "nlevels") and d.columns.nlevels > 1:
            d.columns = d.columns.get_level_values(0)
        return d

    AGG = {"Open": "first", "High": "max", "Low": "min", "Close": "last", "Volume": "sum"}

    results = []
    for t in TARGETS:
        try:
            # 3분봉: yfinance 미지원 → 1분봉을 3분으로 리샘플
            df1 = _flatten(yf.download(t["ticker"], interval="1m", period="5d", progress=False, auto_adjust=True))
            df3 = df1.resample("3min").agg(AGG).dropna() if df1 is not None and len(df1) else df1
            df15 = _flatten(yf.download(t["ticker"], interval="15m", period="1mo", progress=False, auto_adjust=True))
            tf3 = analyze_tf(df3, np)
            tf15 = analyze_tf(df15, np)
            signal, color = classify(tf3)
            entry = {**t, "tf3": tf3, "tf15": tf15, "signal": signal, "color": color}
            results.append(entry)
            if tf3:
                print(f"  {t['key']:10} {signal} | 3m RSI={tf3['rsi7']} z={tf3['z20']} VWAP이격={tf3['vwap_dev']}% (바 {tf3['last_bar']})")
            else:
                print(f"  {t['key']:10} 데이터 부족")
        except Exception as e:
            print(f"  [WARN] {t['key']} 실패: {e}")
            results.append({**t, "tf3": None, "tf15": None, "signal": "수집실패", "color": "muted"})

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "targets": results,
        "disclaimer": "3분봉(1분봉 리샘플)·15분봉. 체결은 HTS/거래소로. yfinance 15~20분 지연 가능 · 단타는 실시간 차트 필수.",
        "note_sk": "SK하이닉스는 yfinance가 정규장(09:00~15:30)만 제공 — 08:00 장전/시간외(~18:00)·20:00 데이터는 HTS에서만 확인. XRP는 24시간.",
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
