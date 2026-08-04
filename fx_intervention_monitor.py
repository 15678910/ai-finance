"""통화 저평가·외환개입 모니터 — REER 괴리 + FIMA 레포 + 미·일 금리차

왜 만들었나
----------
'미 재무부 엔화 개입' 브리핑의 핵심 주장들(엔화·원화 과도한 저평가, 위안화 20~30% 저평가,
미·일 금리차 2.5~2.75%p, FIMA 레포로 개입 자금 조달)은 **전부 숫자로 검증 가능한 명제**다.
읽고 끝내는 대신 매일 자동으로 재는 지표로 바꾼다.

세 축
-----
1) REER(실질실효환율) 괴리 — BIS 광의 실질실효환율(FRED 경유).
   물가차를 반영한 '실질' 통화가치라 명목환율보다 저평가 판단에 적합.
   10년 평균 대비 괴리율과 z-score로 '얼마나 이례적인가'를 잰다.
   ⚠️ IMF·EU가 말하는 저평가율은 '균형환율 모델' 추정치로 방법론이 다르다.
      REER 괴리는 과거 평균 대비 상대적 위치일 뿐 균형가치가 아니다 — 수치가 다를 수 있다.

2) FIMA 레포 잔액 — 외국 중앙은행이 보유 미국채를 연준에 담보로 맡기고 달러를 빌린 잔액.
   시장에 국채를 내다 팔지 않고 개입 실탄을 마련하는 통로라, 잔액 급증은
   **공식 개입이 진행 중이거나 임박했다는 흔적**이 된다. 주간(수요일) 공시.

3) 미·일 정책금리차 — 엔 캐리 트레이드(저금리 엔 차입 → 고금리 자산 투자)의 동력.
   격차가 벌어질수록 엔 약세 압력, 좁아지면 캐리 청산(글로벌 위험자산 동반 매도) 위험.

출력: docs/fx_intervention.json
🚨 공개 데이터의 규칙기반 요약 · 투자자문 아님. 개입 시점·규모를 예측하지 않는다.
"""

import json
import os
import subprocess
import sys
import urllib.parse
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "fx_intervention.json")
FRED_CSV = "https://fred.stlouisfed.org/graph/fredgraph.csv"

# BIS 광의 실질실효환율 (2020=100). 값이 낮을수록 물가 반영 후 통화가 싸다는 뜻.
REER = [("KRW", "RBKRBIS", "원화", "🇰🇷"), ("JPY", "RBJPBIS", "엔화", "🇯🇵"),
        ("CNY", "RBCNBIS", "위안화", "🇨🇳"), ("USD", "RBUSBIS", "달러", "🇺🇸")]
RATE_US = ("DFF", "미 연방기금금리")
RATE_JP = ("IRSTCI01JPM156N", "일본 콜금리")     # INTDSRJPM193N은 2017년에 갱신 중단 → 사용 불가
FIMA = ("WORAL", "FIMA 레포 잔액")               # 단위: 백만 달러, 주간(수)
YEARS = 10


def fred_csv(series_id, start):
    """FRED CSV → [(date, value)]. urllib 우선, 실패 시 curl 폴백.

    일부 개발 환경에서 파이썬 소켓이 FRED에 닿지 못하는(read timeout) 경우가 있어 이중화한다.
    GitHub Actions에서는 urllib 경로로 정상 동작한다.
    """
    url = f"{FRED_CSV}?{urllib.parse.urlencode({'id': series_id, 'cosd': start})}"
    txt = None
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        txt = urllib.request.urlopen(req, timeout=25).read().decode("utf-8", "replace")
    except Exception:
        try:
            txt = subprocess.run(["curl", "-sS", "--max-time", "40", url],
                                 capture_output=True, text=True, timeout=60).stdout
        except Exception as e:
            print(f"  [WARN] {series_id} 수집 실패: {e}")
            return []
    out = []
    for line in (txt or "").strip().split("\n")[1:]:
        parts = line.split(",")
        if len(parts) < 2 or parts[1].strip() in (".", ""):
            continue
        try:
            out.append((parts[0].strip(), float(parts[1])))
        except ValueError:
            continue
    return out


def stats(vals):
    n = len(vals)
    m = sum(vals) / n
    sd = (sum((x - m) ** 2 for x in vals) / n) ** 0.5
    return m, sd


def reer_block():
    """통화별 REER 괴리·z·10년 레인지 내 위치."""
    start = (datetime.now() - timedelta(days=365 * YEARS + 30)).strftime("%Y-%m-%d")
    out = []
    for code, sid, name, flag in REER:
        rows = fred_csv(sid, start)
        if len(rows) < 24:
            print(f"  [WARN] {name}({sid}) 표본 부족 {len(rows)}")
            continue
        vals = [v for _, v in rows]
        cur, asof = vals[-1], rows[-1][0]
        m, sd = stats(vals)
        lo, hi = min(vals), max(vals)
        z = (cur - m) / sd if sd else 0.0
        dev = (cur / m - 1) * 100
        pctile = sum(1 for v in vals if v < cur) / len(vals) * 100
        if z <= -2:
            vd, col = "극단적 저평가", "green"
        elif z <= -1:
            vd, col = "저평가", "cyan"
        elif z >= 2:
            vd, col = "극단적 고평가", "red"
        elif z >= 1:
            vd, col = "고평가", "yellow"
        else:
            vd, col = "중립 범위", "gray"
        out.append({
            "code": code, "name": name, "flag": flag, "series_id": sid, "asof": asof,
            "value": round(cur, 2), "mean10y": round(m, 2),
            "dev_pct": round(dev, 1), "z": round(z, 2),
            "min10y": round(lo, 2), "max10y": round(hi, 2),
            "pctile": round(pctile, 1), "at_10y_low": bool(cur <= lo + 1e-9),
            "verdict": vd, "color": col,
            "spark": [round(v, 2) for v in vals[-60:]],
        })
        print(f"  {flag} {name}: {cur:.2f} (10년평균 {m:.2f} · 괴리 {dev:+.1f}% · z {z:+.2f}) {vd}"
              + (" · 10년 최저" if cur <= lo + 1e-9 else ""))
    return out


def rate_gap_block():
    """미·일 정책금리차 — 엔 캐리 동력."""
    start = (datetime.now() - timedelta(days=800)).strftime("%Y-%m-%d")
    us, jp = fred_csv(RATE_US[0], start), fred_csv(RATE_JP[0], start)
    if not us or not jp:
        return None
    gap = us[-1][1] - jp[-1][1]
    prev_gap = None
    if len(jp) >= 7:                                    # 약 6개월 전 대비 방향
        base_jp = jp[-7]
        base_us = next((v for d, v in reversed(us) if d <= base_jp[0]), None)
        if base_us is not None:
            prev_gap = base_us - base_jp[1]
    trend = None if prev_gap is None else round(gap - prev_gap, 2)
    return {
        "us_rate": round(us[-1][1], 2), "us_asof": us[-1][0], "us_label": RATE_US[1],
        "jp_rate": round(jp[-1][1], 2), "jp_asof": jp[-1][0], "jp_label": RATE_JP[1],
        "gap_pp": round(gap, 2), "gap_chg_6m_pp": trend,
        "carry": "강함" if gap >= 3 else "보통" if gap >= 1.5 else "약함",
        "note": ("금리차가 클수록 엔 차입→고금리 자산 투자(엔 캐리) 유인이 커져 엔 약세 압력. "
                 "반대로 격차가 빠르게 좁아지면 캐리 청산이 몰리며 글로벌 위험자산이 함께 흔들린다."),
    }


def fima_block():
    """FIMA 레포 잔액 — 외국 중앙은행의 달러 조달(개입 실탄) 흔적."""
    rows = fred_csv(FIMA[0], (datetime.now() - timedelta(days=800)).strftime("%Y-%m-%d"))
    if len(rows) < 8:
        return None
    vals = [v for _, v in rows]
    cur, asof = vals[-1], rows[-1][0]
    avg4 = sum(vals[-5:-1]) / 4                          # 직전 4주 평균(당주 제외)
    y52 = vals[-52:] if len(vals) >= 52 else vals
    mx = max(y52)
    # 급증 판정: 절대 규모가 의미 있으면서(≥1,000백만$=$1B) 직전 4주 평균의 3배 이상
    spike = bool(cur >= 1000 and (avg4 <= 0 or cur >= avg4 * 3))
    # 과거 급증 이력 — 지표가 실제로 반응한 적이 있는지 눈으로 확인할 수 있게 함께 싣는다
    hist = sorted([(d, v) for d, v in rows if v >= 1000], key=lambda x: -x[1])[:6]
    spikes = [{"date": d, "musd": round(v, 0), "busd": round(v / 1000, 1),
               "year_end": d[5:7] == "12" and int(d[8:10]) >= 24}
              for d, v in sorted(hist, key=lambda x: x[0], reverse=True)]
    # 다음 공시일 — 매주 수요일 기준, 목요일 공표. '언제 확인하면 되는지'를 알려준다
    try:
        nxt = (datetime.strptime(asof, "%Y-%m-%d") + timedelta(days=7)).strftime("%Y-%m-%d")
    except Exception:
        nxt = None
    return {
        "value_musd": round(cur, 1), "asof": asof, "label": FIMA[1],
        "avg_4w": round(avg4, 1), "max_52w": round(mx, 1),
        "vs_avg4w_x": round(cur / avg4, 1) if avg4 > 0 else None,
        "spike": spike, "next_asof": nxt,
        "past_spikes": spikes,
        "spark": [round(v, 1) for v in vals[-52:]],
        "note": ("각국 중앙은행이 보유 미국채를 연준에 담보로 맡기고 달러를 빌린 잔액(백만$·주간 수요일 기준). "
                 "국채를 시장에 내다 팔지 않고 개입 실탄을 마련하는 통로라, 잔액 급증은 "
                 "공식 개입이 진행 중이거나 임박했다는 흔적으로 읽힌다."),
        "caveat": ("⚠️ 개입 전용 지표가 아니다 — 연말(12월 말) 급증은 자금 결산 수요인 경우가 많고, "
                   "평시 유동성 조달로도 쓰인다. 수요일 스냅샷이라 주 후반 개입은 다음 주에야 잡힌다. "
                   "급증을 보면 '개입이 있었나' 확인의 출발점으로 삼되 단정하지 말 것."),
    }


def build_signal(reers, gap, fima):
    """개입 압력 신호등 — 저평가 심도 + 캐리 유인 + 달러 조달 흔적."""
    score, why = 0, []
    krw = next((r for r in reers if r["code"] == "KRW"), None)
    jpy = next((r for r in reers if r["code"] == "JPY"), None)
    for r in (krw, jpy):
        if not r:
            continue
        if r["z"] <= -2:
            score += 2
            why.append(f"{r['name']} z={r['z']} 극단적 저평가")
        elif r["z"] <= -1:
            score += 1
            why.append(f"{r['name']} z={r['z']} 저평가")
        if r["at_10y_low"]:
            score += 1
            why.append(f"{r['name']} 10년 최저치")
    if gap and gap["gap_pp"] >= 2.5:
        score += 1
        why.append(f"미·일 금리차 {gap['gap_pp']}%p (캐리 유인 {gap['carry']})")
    if fima and fima["spike"]:
        score += 2
        why.append(f"FIMA 레포 급증 {fima['value_musd']}백만$ (4주평균의 {fima['vs_avg4w_x']}배)")
    if score >= 5:
        lvl, col = "🔴 개입 압력 높음", "red"
    elif score >= 3:
        lvl, col = "🟡 개입 압력 누적", "yellow"
    else:
        lvl, col = "🟢 특이 신호 없음", "green"
    return {"score": score, "level": lvl, "color": col, "reasons": why}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    now = datetime.now(KST)
    print("=" * 60)
    print("  통화 저평가·외환개입 모니터 (REER · FIMA 레포 · 미·일 금리차)")
    print("=" * 60)

    print("\n[REER 실질실효환율]")
    reers = reer_block()
    print("\n[미·일 금리차]")
    gap = rate_gap_block()
    if gap:
        print(f"  {gap['us_label']} {gap['us_rate']}% − {gap['jp_label']} {gap['jp_rate']}% "
              f"= {gap['gap_pp']}%p (캐리 {gap['carry']})")
    print("\n[FIMA 레포]")
    fima = fima_block()
    if fima:
        print(f"  {fima['value_musd']}백만$ ({fima['asof']}) · 4주평균 {fima['avg_4w']} · "
              f"52주최대 {fima['max_52w']}" + (" · ⚠️ 급증" if fima["spike"] else ""))

    if not reers:
        print("\n[ERROR] REER 수집 실패 — 기존 파일 보존.")
        return 1

    sig = build_signal(reers, gap, fima)
    print(f"\n[종합] {sig['level']} (점수 {sig['score']})")
    for w in sig["reasons"]:
        print(f"   · {w}")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "signal": sig, "reer": reers, "rate_gap": gap, "fima_repo": fima,
        "method": {
            "reer": "BIS 광의 실질실효환율(2020=100, FRED 경유) · 최근값 vs 10년 평균 괴리·z-score",
            "caveat": ("REER 괴리는 '과거 평균 대비 위치'일 뿐 균형환율이 아니다. "
                       "IMF·EU가 발표하는 저평가율은 균형환율 모델 추정치라 방법론이 달라 "
                       "수치가 크게 다를 수 있다 — 두 값을 같은 것으로 취급하면 안 된다."),
            "fima": "연준 H.4.1 외국 공식기관 레포(백만$, 주간). 개입 외 용도 가능성 있음",
            "rate_gap": "미 연방기금금리(일별) − 일본 콜금리(월별). 기준일이 달라 최대 한 달 시차",
        },
        "portfolio_note": ("원화가 저평가에서 벗어나면(원화 강세) 외국인의 한국 주식 환차익 유인이 커지고, "
                           "반대로 엔 캐리가 청산되면 글로벌 위험자산이 함께 흔들린다 — 두 경로 모두 "
                           "국내 증시 수급에 직접 닿는다. 다만 방향·시점을 예측하는 지표는 아니다."),
        "note": "공개 데이터(FRED/BIS/연준)의 규칙기반 요약 · 투자자문 아님 · 개입 시점·규모를 예측하지 않음",
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"\n[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
