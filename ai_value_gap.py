"""
AI 생태계 밸류에이션 갭 — 발주처(수요) vs 공급처(공급), 한국 반도체 강제 저평가 진단
========================================================================================
논지(사용자 브리핑): AI '토큰 경제'의 병목은 공급처(메모리). 그런데 한국 반도체(삼성·SK)는
전 세계에서 가장 돈 잘 벌면서 가장 낮은 PER(강제 저평가). 본 모듈은 그 PER 격차를 라이브로 검증.

· 발주처(수요): MSFT·GOOGL·META·AMZN·AAPL (하이퍼스케일러)
· 공급처(공급): NVDA·TSM·삼성·SK하이닉스·MU (AI 팩토리 부품)
선행 PER(forwardPE)을 공통 지표로 사용(한국주는 후행 PER 결측 잦음). 시총은 USD 환산.

출력: docs/ai_value_gap.json
🚨 밸류에이션 '사실' + 브리핑 '논지' 분리 표시. 투자자문 아님(라이선스 자문 아님).
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "ai_value_gap.json")

UNIVERSE = [
    ("마이크로소프트", "MSFT", "발주처", "🇺🇸"),
    ("알파벳", "GOOGL", "발주처", "🇺🇸"),
    ("메타", "META", "발주처", "🇺🇸"),
    ("아마존", "AMZN", "발주처", "🇺🇸"),
    ("애플", "AAPL", "발주처", "🇺🇸"),
    ("엔비디아", "NVDA", "공급처", "🇺🇸"),
    ("TSMC", "TSM", "공급처", "🇹🇼"),
    ("삼성전자", "005930.KS", "공급처", "🇰🇷"),
    ("SK하이닉스", "000660.KS", "공급처", "🇰🇷"),
    ("마이크론", "MU", "공급처", "🇺🇸"),
]

# 브리핑 논지(사용자 제공) — 사실 데이터와 분리해 '논지'로 표시
THESIS = {
    "summary": ("AI 산업은 엔비디아 'AI 팩토리'가 토큰(=매출)을 생산하는 구조로 재편. "
                "병목은 수요(빅테크)가 아니라 공급(메모리 반도체). 그런데 한국 반도체(삼성·SK)는 "
                "전 세계에서 가장 돈 잘 벌면서 가장 낮은 PER을 받는 '강제 저평가' 상태."),
    "roadmap": [
        "초기(현재~3년): '토큰 공급자'의 시대 — 공급 병목 핵심인 SK·삼성에 이익 집중.",
        "성숙(3년 이후): AI 승자 SW/서비스(수요처)로 성장 무게중심 이동.",
        "나스닥: capex 부담으로 박스권/조정 가능(부채·유상증자 단행 가능성).",
        "KOSPI: 삼성·SK 이익 안정으로 장기 랠리 + '성장주'로 변모 가능성.",
        "메모리 스페셜티(후공정 패키징) 부상 → SK 위상 강화, SK가 삼성 제치고 시총 1위 역전 가능성.",
    ],
    "conclusion": ("글로벌 조정은 생태계 내 '이익 비중 최대·시총 최저'인 한국 반도체를 저가 매수할 전략적 기회 "
                   "(브리핑 논지). 수급 불안요인 배제 시 펀더멘털상 매력. ※ 논지이며 투자자문 아님."),
    "source": "사용자 제공 브리핑 (AI 생태계 저평가 분석)",
}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    import yfinance as yf
    import warnings
    warnings.filterwarnings("ignore")

    now = datetime.now(KST)

    # USD/KRW 환율(한국 시총 USD 환산용)
    usdkrw = 1380.0
    try:
        fx = yf.Ticker("USDKRW=X").fast_info
        v = float(getattr(fx, "last_price", 0) or 0)
        if v > 500:
            usdkrw = v
    except Exception:
        pass

    items = []
    for nm, tk, grp, flag in UNIVERSE:
        try:
            info = yf.Ticker(tk).info or {}
            tpe = info.get("trailingPE")
            fpe = info.get("forwardPE")
            mc = info.get("marketCap")
            mc_usd = None
            if mc:
                mc_usd = float(mc) / usdkrw if tk.endswith(".KS") else float(mc)
            items.append({
                "name": nm, "ticker": tk, "group": grp, "flag": flag,
                "fwd_pe": round(float(fpe), 1) if fpe else None,
                "trail_pe": round(float(tpe), 1) if tpe else None,
                "mcap_usd_t": round(mc_usd / 1e12, 2) if mc_usd else None,
                "korea": tk.endswith(".KS"),
            })
            print(f"  {flag} {nm:10s} [{grp}] 선행PER {round(float(fpe),1) if fpe else '—'} · 시총 {round(mc_usd/1e12,2) if mc_usd else '—'}T")
        except Exception as e:
            print(f"  [WARN] {nm} 실패: {e}")

    def avg_pe(pred):
        vals = [it["fwd_pe"] for it in items if pred(it) and it["fwd_pe"]]
        return round(sum(vals) / len(vals), 1) if vals else None

    demand_pe = avg_pe(lambda it: it["group"] == "발주처")
    supply_pe = avg_pe(lambda it: it["group"] == "공급처")
    korea_pe = avg_pe(lambda it: it["korea"])
    supply_exkr_pe = avg_pe(lambda it: it["group"] == "공급처" and not it["korea"])
    # 한국 할인율(공급처 비한국 평균 대비)
    discount = None
    if korea_pe and supply_exkr_pe:
        discount = round((1 - korea_pe / supply_exkr_pe) * 100, 1)

    def mcap_sum(pred):
        vals = [it["mcap_usd_t"] for it in items if pred(it) and it["mcap_usd_t"]]
        return round(sum(vals), 2) if vals else None

    print(f"\n  발주처 평균 선행PER {demand_pe} · 공급처 {supply_pe} · 한국 {korea_pe} (비한국 공급처 {supply_exkr_pe}) → 한국 할인율 {discount}%")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "usdkrw": round(usdkrw, 1),
        "items": items,
        "demand_avg_fwd_pe": demand_pe,
        "supply_avg_fwd_pe": supply_pe,
        "korea_avg_fwd_pe": korea_pe,
        "supply_exkorea_fwd_pe": supply_exkr_pe,
        "korea_discount_pct": discount,
        "demand_mcap_t": mcap_sum(lambda it: it["group"] == "발주처"),
        "supply_mcap_t": mcap_sum(lambda it: it["group"] == "공급처"),
        "thesis": THESIS,
        "note": ("선행 PER(forwardPE) 기준 라이브 비교. 한국주는 후행 PER 결측 잦아 선행 사용. "
                 "시총은 USD 환산(한국=원화/USDKRW). '사실=PER 데이터', '논지=브리핑'을 분리 표시. 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
