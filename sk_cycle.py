"""
SK하이닉스 사이클 변곡점 분석 — 주가는 실적·메모리가격을 선행한다
====================================================================
핵심 통찰(과거 변곡점에서 도출):
  · SK 주가는 분기 실적을 1~2분기 '선행'한다.
    - 2021 4Q: 영업이익 정점(4.22조)인데 주가는 이미 고점 → 불황 선반영 하강.
    - 2023 1Q: 적자 저점(-3.40조)인데 주가는 감산 호재로 이미 바닥 반등.
  · 따라서 '현재 실적'이 아니라 '메모리 가격 방향 + 감산/증설 + HBM·AI 수요'가 주가를 끈다.
  · 현재(2026 1Q): 영업이익 37.61조(Record)·마진 72%·100만원 돌파 = AI 황제주 국면.

이 모듈은 변곡점 타임라인(큐레이션)을 보존하고, 현재 주가 위치(52주 고점 근접도·모멘텀)를
계산해 '사이클 국면 + 고점 경계' 컨텍스트를 산출한다. 일간 점 예측이 아니라 '국면/위험' 레이어.

출력: docs/sk_cycle.json
🚨 큐레이션·통계 컨텍스트. 투자 결정 단독 사용 금지.
"""

import json
import os
import re
import sys
import urllib.request
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "sk_cycle.json")
CODE = "000660"

# 주요 변곡점 (사용자 큐레이션 — 분기 영업이익·분기말 종가·국면 특징)
INFLECTIONS = [
    {"q": "2021 4Q", "op_profit": 4.22, "price": 131000, "type": "고점선행",
     "tag": "단기 고점 선행", "desc": "서버 수요로 실적 정점 달성했으나 내년 불황 선반영 하강 준비"},
    {"q": "2023 1Q", "op_profit": -3.40, "price": 88600, "type": "바닥선행",
     "tag": "업황 바닥 선반영", "desc": "적자 폭 최대에도 주가는 감산 호재에 선제 반등 시작"},
    {"q": "2024 2Q", "op_profit": 5.47, "price": 236500, "type": "랠리",
     "tag": "HBM 독점 랠리", "desc": "NVIDIA 공식 독점 지위 부각되며 단기 초고가 달성"},
    {"q": "2025 4Q", "op_profit": 19.17, "price": 720000, "type": "구조성장",
     "tag": "구조적 대세 성장", "desc": "수익 중심 믹스로 분기 영업이익 최고치 경신, 70만닉스 시대"},
    {"q": "2026 1Q", "op_profit": 37.61, "price": 1100000, "type": "정점권",
     "tag": "AI 황제주 등극", "desc": "영업이익률 72% 경이적 마진 창출 및 100만 원 돌파 (Record)"},
]

LESSON = ("SK 주가는 실적·메모리가격을 1~2분기 선행한다 — 2021 4Q 실적정점이 주가고점, "
          "2023 1Q 적자저점이 주가바닥이었다. '현재 실적'이 아니라 '메모리 가격 방향·감산·HBM/AI 수요'가 주가를 끈다.")
WATCH = ["메모리(DRAM·HBM) 고정거래가 하락 전환", "영업이익률 정점 통과(현 72%)",
         "HBM capex 과잉·재고 증가", "마이크론·삼성 증설 발표(공급 확대)"]


def naver_ohlc(code, days=400):
    end = date.today().strftime("%Y%m%d")
    start = (date.today() - timedelta(days=days)).strftime("%Y%m%d")
    url = (f"https://api.finance.naver.com/siseJson.naver?symbol={code}"
           f"&requestType=1&startTime={start}&endTime={end}&timeframe=day")
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0", "Referer": "https://finance.naver.com/"})
    txt = urllib.request.urlopen(req, timeout=15).read().decode("utf-8")
    rows = re.findall(r'\["(\d{8})",\s*[\d.]+,\s*[\d.]+,\s*[\d.]+,\s*([\d.]+)', txt)
    return [(f"{d[:4]}-{d[4:6]}-{d[6:]}", float(c)) for d, c in rows]


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    bars = naver_ohlc(CODE, 400)
    closes = [c for _, c in bars]
    price = closes[-1] if closes else None
    metrics, cycle = {}, {"key": "neutral", "text": ""}

    if price and len(closes) >= 60:
        hi52 = max(closes[-250:]) if len(closes) >= 250 else max(closes)
        lo52 = min(closes[-250:]) if len(closes) >= 250 else min(closes)
        from_hi = (price / hi52 - 1) * 100
        mom3 = (price / closes[-63] - 1) * 100 if len(closes) >= 63 else None
        metrics = {"price": round(price), "week52_high": round(hi52), "week52_low": round(lo52),
                   "pct_from_52wh": round(from_hi, 1), "mom_3m_pct": round(mom3, 1) if mom3 is not None else None}

        # 사이클 국면 판정 — 사상최고 영업이익·100만원 + 고점 근접 → 2021 4Q 패턴 경계
        near_high = from_hi >= -8
        if near_high:
            cycle = {"key": "peak_watch",
                     "text": ("🔴 고점권 경계 — 사상 최고 실적(영업이익률 72%)·주가 100만원대에서 52주 고점 근접. "
                              "2021 4Q처럼 '실적정점=주가고점' 패턴 주의. 단 AI 구조 수요로 과거 사이클과 다를 수 있어 "
                              "메모리 가격 하락 전환이 진짜 분기점.")}
        elif mom3 is not None and mom3 <= -15:
            cycle = {"key": "correction",
                     "text": "🟡 조정권 — 고점 대비 큰 폭 하락. 2023 1Q처럼 '감산·가격반등' 신호가 바닥 선행 단서."}
        else:
            cycle = {"key": "upcycle",
                     "text": "🟢 상승 사이클 진행 — 메모리 가격·HBM 수요 우호. 고점 경계선까지 여유."}
        print(f"  SK 현재가 {price:,.0f} · 52주고점 대비 {from_hi:+.1f}% · 3개월 {mom3:+.1f}% → {cycle['key']}")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "name": "SK하이닉스",
        # 기록: 2026-06-23 폭락 (peak_watch 경고의 현실화)
        "recent_shock": {
            "date": "2026-06-23",
            "title": "KOSPI -10% 폭락 (장중 서킷브레이커)",
            "text": ("SK하이닉스 -12.5%·삼성전자 -12.3% 동시 급락. 원인=ADR 유상증자(신주발행) 희석 우려 + "
                     "SK 연초 대비 +348% 극단 쏠림(6/22 삼성 제치고 시총 1위)의 되돌림 — 반도체 고유 사건. "
                     "레버리지 ETF·변동성 디레버리징·서킷브레이커는 '원인'이 아니라 '증폭기'. "
                     "sk_cycle '고점권 경계'와 ADR '유상증자 오버행' 경고가 현실화된 사례."),
        },
        "current_phase": INFLECTIONS[-1]["tag"],
        "metrics": metrics,
        "cycle": cycle,
        "inflections": INFLECTIONS,
        "lesson": LESSON,
        "watch": WATCH,
        "note": ("변곡점은 큐레이션(분기 영업이익·종가). 현재 국면은 주가 위치(52주 고점 근접·모멘텀)로 산출. "
                 "주가가 실적을 선행하므로 메모리 가격·감산·HBM 뉴스가 핵심 — 긴급경보(메모리·SK실적)와 연동."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
