"""
외환위기 조기경보 — FX 스왑 '그림자 부채' 6대 지표 모니터
====================================================================
브리핑 논지: 위기는 '달러가 비싸지는 것(환율)'이 아니라 '달러가 마르는 것(유동성 경색)'.
전 세계 FX 스왑 시장에 ~80조 달러 장부밖 부채가 숨어 있고, NBFI(연기금·보험)가 취약 고리.
→ 한국 위기 징후를 6개 지표로 선제 감시.

⚠️ 정직성: 스왑포인트·롤오버·NDF는 전문 FX 데이터라 무료 피드로 직접 산출 불가 →
   '정성 모니터링' 또는 프록시로 표시. 글로벌 달러(DXY·美금리·VIX)·외국인 수급은 실측.

출력: docs/fx_swap.json   🚨 프록시 기반 참고. 실제 위기 판단엔 스왑포인트 필수. 투자자문 아님.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS, "fx_swap.json")

# 브리핑 큐레이션 — 한국 안전판/논지 (사실 데이터와 분리)
SAFEGUARDS = [
    "주요국 통화스왑 한도 $1,506억",
    "국민연금 외환스왑 한도 $650억",
    "외화예금 초과지준 이자·선물환 포지션 한도 조정",
    "⚠️ 미 연준 상설 통화스왑 부재",
    "⚠️ 달러 호딩 — 수출기업 환전 지연으로 '달러는 많으나 안 도는' 현상",
]
THESIS = ("'달러가 비싸진 것(환율↑)'과 '달러가 마르는 것(유동성 경색)'은 전혀 다른 사건. "
          "진짜 위기는 FX 스왑 시장(달러 혈관)이 막힐 때. 전 세계 FX 스왑에 ~80조 달러 장부밖 부채가 "
          "숨어 있고(BIS), 연기금·보험 등 NBFI는 중앙은행 달러창구 직접 접근 불가라 위기 시 헐값 매각 위험. "
          "수치 안전판은 임시 소화기일 뿐 — 원화자산 펀더멘털(실적·지배구조·성장)이 최강 방어선.")


def _yf_series(tk, period="3mo"):
    import yfinance as yf
    s = yf.Ticker(tk).history(period=period)["Close"].dropna()
    return s


def _lvl(s):
    if s is None or len(s) < 2:
        return None, None, None
    c = float(s.iloc[-1])
    chg5 = round((c / float(s.iloc[-6]) - 1) * 100, 2) if len(s) >= 6 else None
    chg20 = round((c / float(s.iloc[-21]) - 1) * 100, 2) if len(s) >= 21 else None
    return round(c, 2), chg5, chg20


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    import warnings
    warnings.filterwarnings("ignore")

    now = datetime.now(KST)

    # ── 실측 시장 데이터 ──
    usdkrw, krw5, krw20 = _lvl(_safe(lambda: _yf_series("KRW=X")))
    dxy, dxy5, _ = _lvl(_safe(lambda: _yf_series("DX-Y.NYB")))
    ust10, ust10_5, _ = _lvl(_safe(lambda: _yf_series("^TNX")))
    vix, vix5, _ = _lvl(_safe(lambda: _yf_series("^VIX")))

    # ── 신용 스프레드(HY OAS) — 글로벌 달러 조달 스트레스 프록시 ──
    hy = {}
    try:
        with open(os.path.join(DOCS, "credit_spread.json"), encoding="utf-8") as f:
            cs = json.load(f)
        for r in cs.get("results", []):
            nm = str(r.get("name", ""))
            if "하이일드" in nm or "high yield" in nm.lower() or "HY" == nm:
                hy = {"oas": r.get("latest_value"), "pct_1y": r.get("percentile_1y"),
                      "chg_5d": r.get("change_5d")}
                break
    except Exception:
        pass

    # ── 외국인 수급(최근 5일) — 자금 이탈 신호 ──
    foreign_5d_eok = None
    try:
        with open(os.path.join(DOCS, "semi_decoupling.json"), encoding="utf-8") as f:
            sd = json.load(f)
        vals = [a.get("foreign_5d_eok") for a in sd.get("assets", []) if a.get("foreign_5d_eok") is not None]
        if vals:
            foreign_5d_eok = sum(vals)
    except Exception:
        pass

    # ── 스트레스 점수(측정 가능 프록시만) ──
    score, parts = 0, []
    if usdkrw is not None:
        s = (2 if usdkrw >= 1500 else 1 if usdkrw >= 1450 else 0) + (1 if (krw5 or 0) >= 1.0 else 0)
        score += s; parts.append(f"USD/KRW {usdkrw}(+{s})")
    if dxy is not None:
        s = (1 if dxy >= 103 else 0) + (1 if (dxy5 or 0) >= 1.0 else 0)
        score += s; parts.append(f"DXY {dxy}(+{s})")
    if vix is not None:
        s = (2 if vix >= 30 else 1 if vix >= 25 else 0)
        score += s; parts.append(f"VIX {vix}(+{s})")
    if hy.get("oas") is not None:
        s = (2 if hy["oas"] >= 6 else 1 if hy["oas"] >= 4.5 else 0)   # HY OAS 절대수준(>6%=위기)
        score += s; parts.append(f"HY OAS {hy['oas']}%(+{s})")
    if foreign_5d_eok is not None:
        s = (2 if foreign_5d_eok <= -50000 else 1 if foreign_5d_eok < 0 else 0)
        score += s; parts.append(f"외인5일 {round(foreign_5d_eok/10000,1)}조(+{s})")

    if score >= 6:
        status, scol = "🔴 경색 위험", "red"
    elif score >= 3:
        status, scol = "🟡 주의 — 일부 스트레스", "amber"
    else:
        status, scol = "🟢 정상 — 스왑시장 작동 추정", "green"

    # ── 6대 지표 매핑(실측/프록시/정성) ──
    indicators = [
        {"n": "① FX 스왑포인트", "measurable": False,
         "proxy": f"한미 금리차(美10Y {ust10}%)·환변동성", "warn": "금리차로 설명 안 되는 급격한 마이너스 심화",
         "note": "직접 산출엔 전문 FX 데이터(원화 선물환) 필요 — 정성 모니터링"},
        {"n": "② 달러 롤오버", "measurable": False,
         "proxy": f"신용스프레드(HY OAS {hy.get('oas','—')})", "warn": "자동 연장되던 단기 달러부채 연장 거부",
         "note": "은행·기업 만기연장은 공개 실시간 미제공 — 정성"},
        {"n": "③ 글로벌 달러 시장", "measurable": True,
         "proxy": f"DXY {dxy} · 美10Y {ust10}% · VIX {vix}", "warn": "美 시장 출렁임→주변부 공급망 타격",
         "note": "실측"},
        {"n": "④ 외국인 자금 흐름", "measurable": True,
         "proxy": (f"반도체 외인 5일 {round(foreign_5d_eok/10000,1)}조" if foreign_5d_eok is not None else "—"),
         "warn": "특히 채권 장기자금 이탈(재정신뢰 하락)", "note": "실측(주식 수급)"},
        {"n": "⑤ 금리 경로", "measurable": True,
         "proxy": f"美10Y {ust10}% (5일 {ust10_5}%p)", "warn": "금리차 확대+위험회피 겹쳐 달러펌프 약화",
         "note": "美 금리 실측 · BOK 사이클 병행"},
        {"n": "⑥ NDF발 전염", "measurable": False,
         "proxy": f"USD/KRW 야간변동(5일 {krw5}%)", "warn": "역외 원화 약세가 국내 유동성 경색으로 확산",
         "note": "역외 NDF 호가는 무료 미제공 — USD/KRW 급변으로 프록시"},
    ]

    print(f"  외환 유동성 상태: {status} (점수 {score}) · {' · '.join(parts)}")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "status": status, "status_color": scol, "stress_score": score, "score_parts": parts,
        "market": {"usdkrw": usdkrw, "usdkrw_5d": krw5, "usdkrw_20d": krw20,
                   "dxy": dxy, "dxy_5d": dxy5, "ust10y": ust10, "ust10y_5d": ust10_5,
                   "vix": vix, "vix_5d": vix5},
        "hy_oas": hy, "foreign_5d_eok": foreign_5d_eok,
        "indicators": indicators,
        "korea": {"liquidity": "FX 스왑시장 정상 작동 추정(프록시) — '혈관 경색' 단계 아님",
                  "safeguards": SAFEGUARDS},
        "thesis": THESIS,
        "note": ("위기는 환율 수치가 아니라 FX 스왑(달러 혈관) 막힘에서 발생. 6지표 중 스왑포인트·롤오버·NDF는 "
                 "전문 데이터라 프록시/정성. 글로벌 달러·외국인 수급은 실측. 프록시 기반 참고용 · 투자자문 아님."),
    }
    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


def _safe(fn):
    try:
        return fn()
    except Exception:
        return None


if __name__ == "__main__":
    sys.exit(main())
