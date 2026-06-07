"""
시장감정지수 (Fear & Greed) — 자체 composite
=============================================
공식 한국 공포탐욕지수가 없어, 보유 데이터로 직접 계산.
구성: 변동성(VIX) · 모멘텀(z-score) · 자금흐름(외국인) · 포지셔닝(레버리지) · 리스크(신용 스프레드)
0=극단공포(역발상 매수) ↔ 100=극단탐욕(과열·차익 경계)

암호화폐 공포탐욕은 프론트엔드에서 alternative.me 실시간 호출.
출력: docs/sentiment.json
🚨 통계 추정. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS, "sentiment.json")


def _load(name):
    try:
        with open(os.path.join(DOCS, name), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _clamp(v, lo=0.0, hi=100.0):
    return max(lo, min(hi, v))


def vix_to_score(v):
    # 낮은 VIX=탐욕(안일), 높은 VIX=공포
    if v is None:
        return None
    return round(_clamp(100 - (v - 12) * 3.3), 1)


def z_to_score(z):
    if z is None:
        return None
    return round(_clamp(50 + z * 22), 1)


def credit_to_score(color):
    return {"green": 72, "amber": 50, "yellow": 50, "red": 28}.get(color, 50)


def lev_to_score(r):
    if r is None:
        return None
    return round(_clamp(50 + (r - 1) * 40), 1)


def flow_to_score(net_shares):
    # 46개 종목 5일 외국인 순매수 합산 → 임계 1.5억주로 정규화
    if net_shares is None:
        return None
    import math
    return round(_clamp(50 + 50 * math.tanh(net_shares / 1.5e8)), 1)


def classify(score):
    if score >= 75:
        return "극단적 탐욕", "red", "과열·차익 경계 (역발상: 분할 차익)"
    if score >= 60:
        return "탐욕", "orange", "위험선호 강함 — 추격 주의"
    if score >= 45:
        return "중립", "amber", "방향성 모호 — 관망"
    if score >= 25:
        return "공포", "cyan", "위험회피 — 분할매수 관심"
    return "극단적 공포", "green", "투매 국면 (역발상: 매수 유리 구간)"


def weighted(components):
    """components: [(name, raw_str, score, weight)] — score None은 제외하고 가중평균."""
    valid = [(s, w) for (_, _, s, w) in components if s is not None]
    if not valid:
        return None
    tw = sum(w for _, w in valid)
    return round(sum(s * w for s, w in valid) / tw, 1)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    print("=" * 55)
    print("  시장감정지수 (자체 composite)")
    print(f"  KST: {datetime.now(KST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    om = _load("overseas_market.json")
    oh = _load("overheating.json")
    lv = _load("leverage_volatility.json")
    data = _load("data.json")

    # VIX
    vix = None
    for m in om.get("all_markets", []):
        if m.get("ticker") == "^VIX" or m.get("is_vix"):
            vix = m.get("current")
            break

    # 모멘텀 z (overheating indices)
    zmap = {i["key"]: i.get("z50") for i in oh.get("indices", [])}
    us_z = [zmap.get("SOX"), zmap.get("NDX")]
    kr_z = [zmap.get("KOSPI"), zmap.get("EWY")]
    us_z = [z for z in us_z if z is not None]
    kr_z = [z for z in kr_z if z is not None]
    us_z_avg = sum(us_z) / len(us_z) if us_z else None
    kr_z_avg = sum(kr_z) / len(kr_z) if kr_z else None

    # 신용 regime
    credit_color = (data.get("credit_spread") or {}).get("macro_regime_color")

    # 레버리지 (국내)
    lev_ratio = lv.get("lev_inv_ratio")

    # 외국인 자금흐름 (investor_flow results 합산)
    foreign_net = None
    fl = data.get("investor_flow") or {}
    results = fl.get("results") or []
    if results:
        tot = 0
        cnt = 0
        for r in results:
            for k in ("foreign_5d_net_shares", "foreign_net_shares", "foreign_10d_net_shares", "foreign_net"):
                if isinstance(r.get(k), (int, float)):
                    tot += r[k]
                    cnt += 1
                    break
        foreign_net = tot if cnt else None

    # ── 미국 감정지수 ──
    us_comps = [
        ("변동성(VIX)", f"{vix}", vix_to_score(vix), 0.40),
        ("모멘텀(z)", f"{us_z_avg:+.2f}σ" if us_z_avg is not None else "—", z_to_score(us_z_avg), 0.35),
        ("신용 리스크", str(credit_color or "—"), credit_to_score(credit_color), 0.25),
    ]
    us_score = weighted(us_comps)

    # ── 한국 감정지수 ──
    kr_comps = [
        ("변동성(VIX)", f"{vix}", vix_to_score(vix), 0.25),
        ("모멘텀(z)", f"{kr_z_avg:+.2f}σ" if kr_z_avg is not None else "—", z_to_score(kr_z_avg), 0.30),
        ("외국인 자금", f"{(foreign_net/1000):+.0f}K주" if foreign_net is not None else "—", flow_to_score(foreign_net), 0.25),
        ("레버리지", f"{lev_ratio}" if lev_ratio is not None else "—", lev_to_score(lev_ratio), 0.20),
    ]
    kr_score = weighted(kr_comps)

    def pack(score, comps):
        if score is None:
            return {"score": None, "label": "데이터부족", "color": "muted", "hint": "", "components": []}
        label, color, hint = classify(score)
        return {
            "score": score, "label": label, "color": color, "hint": hint,
            "components": [{"name": n, "raw": r, "score": s, "weight": w} for (n, r, s, w) in comps],
        }

    us = pack(us_score, us_comps)
    kr = pack(kr_score, kr_comps)
    print(f"  🇺🇸 미국: {us['score']} {us['label']}  (VIX={vix}, z={us_z_avg}, credit={credit_color})")
    print(f"  🇰🇷 한국: {kr['score']} {kr['label']}  (z={kr_z_avg}, 외인={foreign_net}, lev={lev_ratio})")

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "us": us, "kr": kr,
        "crypto_source": "alternative.me",  # 프론트 실시간
        "note": "변동성·모멘텀·자금흐름·포지셔닝·신용 결합. 0=극단공포(역발상 매수)↔100=극단탐욕(차익). 통계 추정.",
    }
    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
