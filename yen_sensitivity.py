"""
엔/달러 → 한국 증시 민감도 (레짐별 · 섹터별)
===============================================
핵심 발견: 엔-한국 관계는 레짐에 따라 '부호가 뒤집힌다'.
  · 평온일(|Δ|<1%): 약한 음(-) — 수출경쟁 채널(엔약세=한국 수출 불리), 거의 노이즈(R²~0)
  · 엔 급변일(|Δ|>=1%): 양(+)으로 전환 — 리스크온/오프·캐리 채널 지배
      (엔강세=USDJPY↓=캐리청산·리스크오프=한국 하락 / 엔약세=리스크온=한국 상승)
  · SK하이닉스가 가장 민감(대급변 R²~0.38) — 2024.8 엔캐리 청산 메커니즘.

→ 단일 베타는 레짐을 섞어 오해를 부르므로 '평온일/급변일'로 쪼개 제공. 동시상관·예측 아님.

출력: docs/yen_sensitivity.json
"""

import json
import os
import re
import sys
import urllib.request
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "yen_sensitivity.json")
ASSETS = [("KOSPI", "KOSPI"), ("SK하이닉스", "000660"), ("삼성전자", "005930")]


def naver_close(sym, days=760):
    end = date.today().strftime("%Y%m%d")
    start = (date.today() - timedelta(days=days)).strftime("%Y%m%d")
    url = (f"https://api.finance.naver.com/siseJson.naver?symbol={sym}"
           f"&requestType=1&startTime={start}&endTime={end}&timeframe=day")
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0", "Referer": "https://finance.naver.com/"})
    txt = urllib.request.urlopen(req, timeout=15).read().decode("utf-8")
    rows = re.findall(r'\["(\d{8})",\s*[\d.]+,\s*[\d.]+,\s*[\d.]+,\s*([\d.]+)', txt)
    return {f"{d[:4]}-{d[4:6]}-{d[6:]}": float(c) for d, c in rows}


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

    jpy = yf.download("JPY=X", period="3y", interval="1d", progress=False)["Close"]
    if hasattr(jpy, "columns"):
        jpy = jpy.iloc[:, 0]
    jpy = jpy.dropna()
    jr = {d.strftime("%Y-%m-%d"): float(v) for d, v in jpy.pct_change().items() if not np.isnan(v)}

    def reg(pairs):
        if len(pairs) < 10:
            return {"beta": None, "r": None, "n": len(pairs)}
        x = np.array([p[0] for p in pairs])
        y = np.array([p[1] for p in pairs])
        b = float(np.polyfit(x, y, 1)[0])
        r = float(np.corrcoef(x, y)[0, 1])
        return {"beta": round(b, 2), "r": round(r, 2), "r2": round(r * r, 2), "n": len(pairs)}

    out_assets = []
    for nm, sym in ASSETS:
        try:
            px = naver_close(sym)
        except Exception as e:
            print(f"  [WARN] {nm} 수집 실패: {e}")
            continue
        pk = sorted(px)
        pairs = []
        for i in range(1, len(pk)):
            D, P = pk[i], pk[i - 1]
            if D in jr and px[P] > 0:
                pairs.append((jr[D], px[D] / px[P] - 1))
        quiet = [p for p in pairs if abs(p[0]) < 0.01]
        vol = [p for p in pairs if abs(p[0]) >= 0.01]
        wk = [p[1] for p in pairs if p[0] > 0.003]    # 엔 약세일(USD/JPY +0.3%↑)
        stg = [p[1] for p in pairs if p[0] < -0.003]  # 엔 강세일
        out_assets.append({
            "name": nm, "code": sym,
            "all": reg(pairs), "quiet": reg(quiet), "vol": reg(vol),
            "weak_mean": round(float(np.mean(wk)) * 100, 2) if len(wk) >= 10 else None, "weak_n": len(wk),
            "strong_mean": round(float(np.mean(stg)) * 100, 2) if len(stg) >= 10 else None, "strong_n": len(stg),
        })
        rv = reg(vol)
        print(f"  {nm}: 평온 {reg(quiet)['beta']} / 급변 {rv['beta']} (R²{rv.get('r2')}, n{rv['n']})")

    # 현재 USD/JPY 레벨 + 개입·인상 경계 (160 돌파 = 일본 당국 개입 경계선)
    usdjpy = float(jpy.iloc[-1])
    if usdjpy >= 160:
        ivl, imsg = "경계", ("160 돌파 — 일본 당국 개입·BOJ 인상 트리거 구간. "
                             "약세 자체보다 '급반전(개입·인상)→엔캐리 청산→리스크오프'가 반도체 단기 충격 위험.")
    elif usdjpy >= 157:
        ivl, imsg = "관찰", "157+ — 개입 경계선 접근. 추가 약세 시 경계."
    else:
        ivl, imsg = "정상", ""

    out = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "window_desc": "약 3년 일별 (동시상관)",
        "usdjpy": round(usdjpy, 2),
        "intervention": {"level": ivl, "msg": imsg, "threshold": 160},
        "assets": out_assets,
        "insight": ("평온일(±1% 미만)엔 엔↔한국 거의 무관(약한 음의 수출경쟁). "
                    "엔 급변일(±1% 이상)엔 부호가 양(+)으로 뒤집혀 리스크온/오프 채널 지배 — "
                    "엔강세(USDJPY↓)=캐리청산·위험신호, 엔약세=리스크온. SK하이닉스가 가장 민감."),
        "note": "동시상관·레짐 의존(예측 아님). +1% = USD/JPY 1% 상승(엔 약세) 당 종목 수익률.",
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
