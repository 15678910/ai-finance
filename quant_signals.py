"""
퀀트 신호 생성기 (월드퀀트 우승자 인터뷰 기반)
================================================
세계 퀀트 대회 우승자 김민겸(UNIST) 인터뷰의 핵심 전략 3종을 구현:

  ① 이름 유사도 이상급등 탐지 (텍스트 유사도 알파)
     - 무관한 종목이 이름 유사성만으로 동반 급등하는 비이성적 반응 포착
  ② 롱숏 시장중립 스코어 (Long-Short Equity)
     - PER·ROE·레짐·센티멘트 결합 → 종목별 롱/숏 점수, 시장방향 무관
  ③ 생존 편향 경고 (Survivorship Bias)
     - 현재 생존 종목만 다루는 데이터에 백테스트 부적합 라벨

출력: docs/quant_signals.json

🚨 시뮬레이션·분석용. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
import re
from datetime import datetime, timezone, timedelta
from difflib import SequenceMatcher

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "docs", "data.json")
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "quant_signals.json")


def load_stocks() -> list:
    """data.json에서 전 종목 평탄화."""
    try:
        with open(DATA_FILE, encoding="utf-8") as f:
            d = json.load(f)
    except Exception as e:
        print(f"[ERROR] data.json 로드 실패: {e}")
        return []
    stocks = []
    for sec in d.get("sectors", []):
        for st in sec.get("stocks", []):
            if st.get("ticker"):
                stocks.append({**st, "sector": sec.get("name", "")})
    return stocks


# ─────────────────────────────────────────────────────────────────────
# ① 이름 유사도 이상급등 탐지
# ─────────────────────────────────────────────────────────────────────
# 알려진 재벌 그룹 식별자 (같은 그룹이면 동조가 '합리적')
CHAEBOL_GROUPS = {
    "삼성": ["삼성"], "LG": ["LG"], "SK": ["SK"], "현대": ["현대"],
    "한화": ["한화"], "롯데": ["롯데"], "포스코": ["POSCO", "포스코"],
    "두산": ["두산"], "에코프로": ["에코프로"], "HD현대": ["HD현대", "HD한국"],
    "카카오": ["카카오"], "네이버": ["네이버", "NAVER"],
}


def _group_of(name: str) -> str:
    """기업명이 속한 재벌 그룹 반환 (없으면 '')."""
    for grp, keys in CHAEBOL_GROUPS.items():
        if any(name.startswith(k) for k in keys):
            return grp
    return ""


def _name_tokens(name: str) -> str:
    """비교용: 특수문자만 제거, 핵심 토큰은 유지 (과도 제거 방지)."""
    return re.sub(r"[^가-힣A-Za-z0-9]", "", name or "")


def detect_name_similarity(stocks: list) -> list:
    """이름 유사 + 동반 급등 패턴 탐지 → 관계사/무관 분류.

    퀀트 알파 원리(보고서): 화제 종목과 이름만 비슷한 *무관* 종목이
    비이성적으로 동반 급등할 때 괴리를 알파로 전환.
    같은 재벌 계열사의 동조는 '합리적'이므로 의심도를 낮춘다.
    """
    results = []
    n = len(stocks)
    for i in range(n):
        a = stocks[i]
        na = _name_tokens(a["name"])
        if len(na) < 2:
            continue
        chg_a = a.get("change_pct") or 0
        ga = _group_of(a["name"])
        for j in range(i + 1, n):
            b = stocks[j]
            nb = _name_tokens(b["name"])
            if len(nb) < 2:
                continue
            sim = SequenceMatcher(None, na, nb).ratio()
            if na in nb or nb in na:
                sim = max(sim, 0.72)
            if sim < 0.55:
                continue
            chg_b = b.get("change_pct") or 0
            gb = _group_of(b["name"])
            same_group = bool(ga) and ga == gb     # 같은 재벌 = 합리적
            both_up   = chg_a >= 2.5 and chg_b >= 2.5
            both_down = chg_a <= -2.5 and chg_b <= -2.5
            co_move = abs(chg_a) + abs(chg_b) if (both_up or both_down) else 0
            if co_move == 0:
                continue
            # 의심도: 무관 종목일수록 ↑ (비이성), 같은 그룹이면 대폭 ↓
            rel_mult = 0.25 if same_group else 1.4
            diff_sector = a["sector"] != b["sector"]
            suspicion = round(sim * 100 * rel_mult * (1.15 if diff_sector else 1.0)
                              * min(co_move / 10, 1.6), 1)
            results.append({
                "stock_a": {"name": a["name"], "ticker": a["ticker"], "sector": a["sector"], "change_pct": round(chg_a, 2)},
                "stock_b": {"name": b["name"], "ticker": b["ticker"], "sector": b["sector"], "change_pct": round(chg_b, 2)},
                "similarity": round(sim * 100, 1),
                "diff_sector": diff_sector,
                "same_group": same_group,
                "group": ga if same_group else "",
                "relation": "관계사(합리적 동조)" if same_group else "무관(비이성 의심)",
                "co_move": "동반급등" if both_up else "동반급락",
                "suspicion": suspicion,
            })
    # 무관(비이성) 우선, 그 안에서 의심도순
    results.sort(key=lambda x: (not x["same_group"], x["suspicion"]), reverse=True)
    return results[:12]


# ─────────────────────────────────────────────────────────────────────
# ② 롱숏 시장중립 스코어
# ─────────────────────────────────────────────────────────────────────
def _safe_float(v):
    try:
        if v in (None, "N/A", ""):
            return None
        return float(v)
    except (ValueError, TypeError):
        return None


def compute_long_short(stocks: list) -> dict:
    """PER·ROE·성장·레짐·센티멘트 결합 → 롱/숏 스코어.

    시장중립 원리: 시장 전체 방향과 무관하게, 펀더멘털·모멘텀이
    상대적으로 우수한 종목 Long / 열위 종목 Short.
    각 팩터를 백분위(percentile)로 정규화해 합산.
    """
    REGIME_SCORE = {"강한상승": 1.0, "상승": 0.6, "약한상승": 0.3, "횡보": 0.0,
                    "약한하락": -0.3, "하락": -0.6, "강한하락": -1.0}

    # 팩터 수집
    pool = []
    for st in stocks:
        per = _safe_float(st.get("per") or st.get("forward_per"))
        roe = _safe_float(st.get("roe"))
        growth = _safe_float(st.get("revenue_growth"))
        regime = REGIME_SCORE.get(st.get("regime", ""), None)
        rec_mean = _safe_float(st.get("recommendation_mean"))  # 1=강매수, 5=강매도
        if per is None and roe is None and regime is None:
            continue
        pool.append({
            "name": st["name"], "ticker": st["ticker"], "sector": st["sector"],
            "per": per, "roe": roe, "growth": growth,
            "regime": regime, "rec_mean": rec_mean,
            "change_pct": _safe_float(st.get("change_pct")) or 0,
        })

    # 백분위 정규화 헬퍼
    def percentile_rank(vals: list, v, invert=False):
        valid = [x for x in vals if x is not None]
        if not valid or v is None:
            return 0.5
        below = sum(1 for x in valid if x < v)
        p = below / len(valid)
        return (1 - p) if invert else p

    per_vals = [p["per"] for p in pool]
    roe_vals = [p["roe"] for p in pool]
    growth_vals = [p["growth"] for p in pool]
    rec_vals = [p["rec_mean"] for p in pool]

    for p in pool:
        # 가치(저PER 우수=invert), 수익성(고ROE), 성장(고성장),
        # 모멘텀(레짐), 컨센서스(저rec_mean=매수)
        f_value  = percentile_rank(per_vals, p["per"], invert=True)   # PER 낮을수록 좋음
        f_quality= percentile_rank(roe_vals, p["roe"])
        f_growth = percentile_rank(growth_vals, p["growth"])
        f_mom    = (p["regime"] + 1) / 2 if p["regime"] is not None else 0.5
        f_cons   = percentile_rank(rec_vals, p["rec_mean"], invert=True) if p["rec_mean"] else 0.5
        # 가중합 (0~1) → -100~+100 스코어
        composite = (f_value * 0.25 + f_quality * 0.25 + f_growth * 0.15
                     + f_mom * 0.20 + f_cons * 0.15)
        p["score"] = round((composite - 0.5) * 200, 1)
        p["factors"] = {
            "가치": round(f_value * 100),
            "수익성": round(f_quality * 100),
            "성장": round(f_growth * 100),
            "모멘텀": round(f_mom * 100),
            "컨센서스": round(f_cons * 100),
        }

    pool.sort(key=lambda x: x["score"], reverse=True)
    longs = pool[:8]
    shorts = pool[-8:][::-1]
    return {
        "long": [{"name": p["name"], "ticker": p["ticker"], "sector": p["sector"],
                  "score": p["score"], "factors": p["factors"], "change_pct": p["change_pct"]} for p in longs],
        "short": [{"name": p["name"], "ticker": p["ticker"], "sector": p["sector"],
                   "score": p["score"], "factors": p["factors"], "change_pct": p["change_pct"]} for p in shorts],
        "universe": len(pool),
    }


# ─────────────────────────────────────────────────────────────────────
# ③ 생존 편향 경고 메타데이터
# ─────────────────────────────────────────────────────────────────────
def survivorship_warning() -> dict:
    return {
        "title": "생존 편향 (Survivorship Bias) 주의",
        "body": (
            "본 대시보드의 시가총액 순위·ETF 구성·트리맵·롱숏 스코어는 모두 "
            "'현재 생존 종목'만 다룹니다. 이 데이터로 과거 수익률을 백테스트하면 "
            "상장폐지·퇴출 종목이 빠져 수익률이 비정상적으로 높게 왜곡됩니다."
        ),
        "rules": [
            "현재 종목 리스트로 백테스트 금지 (상장폐지 종목 포함 데이터셋 필요)",
            "'M7 대상 백테스트 로직' 등 SNS 판매 상품은 전형적 생존 편향 결과물",
            "Look-ahead Bias: 미래 정보를 과거 결정에 반영하지 말 것",
            "Data Snooping: 원하는 결과 나올 때까지 데이터 반복 고문 = 과적합",
        ],
        "source": "월드퀀트 우승자 김민겸(UNIST) 인터뷰",
    }


def main():
    print("=" * 55)
    print("  퀀트 신호 생성 (월드퀀트 전략 3종)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    stocks = load_stocks()
    if not stocks:
        print("[ERROR] 종목 데이터 없음")
        return 1
    print(f"\n[로드] {len(stocks)}개 종목")

    print("\n[①] 이름 유사도 이상급등 탐지...")
    name_sim = detect_name_similarity(stocks)
    print(f"  → {len(name_sim)}건 탐지")
    for r in name_sim[:3]:
        print(f"    {r['stock_a']['name']} ↔ {r['stock_b']['name']} "
              f"(유사도 {r['similarity']}%, 의심도 {r['suspicion']})")

    print("\n[②] 롱숏 시장중립 스코어...")
    ls = compute_long_short(stocks)
    print(f"  → 유니버스 {ls['universe']}개")
    if ls["long"]:
        print(f"    LONG 1위: {ls['long'][0]['name']} ({ls['long'][0]['score']:+.1f})")
    if ls["short"]:
        print(f"    SHORT 1위: {ls['short'][0]['name']} ({ls['short'][0]['score']:+.1f})")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "name_similarity": name_sim,
        "long_short": ls,
        "survivorship": survivorship_warning(),
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
