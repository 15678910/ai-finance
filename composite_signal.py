"""
종합 예측 신호 보드 — 다요소 방향 종합 + 디커플링 경고
==========================================================
개별 지표(선물·VKOSPI·과열·디커플링·외인수급·감정·엔·유동성·FX·지정학·M2·이벤트)를
읽어 요소별 방향(vote −2..+2)·가중치로 종합 bias를 산출.
핵심: '선물모델 방향'과 '종합 방향'이 어긋나면 ⚠️ 디커플링 경고 → 오늘(7/1)처럼
선물은 상승인데 실제 급락하는 날을 사전 경고.

출력: docs/composite_signal.json
🚨 통계·지표 종합. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS, "composite_signal.json")


def _load(name):
    try:
        with open(os.path.join(DOCS, f"{name}.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _clamp(v, lo=-2, hi=2):
    return max(lo, min(hi, v))


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    ks = _load("kospi_scenario")
    vk = _load("vkospi")
    oh = _load("overheating")
    sd = _load("semi_decoupling")
    se = _load("sentiment")
    fx = _load("fx_swap")
    lq = _load("liquidity_stress")
    m2 = _load("m2_data")
    dj = _load("data")
    ke = _load("key_events")

    factors = []   # {name, value, signal, vote, weight}

    def add(name, value, signal, vote, weight):
        factors.append({"name": name, "value": value, "signal": signal,
                        "vote": round(vote, 2), "weight": weight})

    # 1) 나스닥 선물(24h) — 1순위 방향 동력
    fut = (ks.get("futures") or {})
    fp = fut.get("fut_overnight_pct")
    if fp is not None:
        vote = _clamp(fp)  # ±1%p ≈ ±1표
        sig = "강한 상승" if fp > 1 else "상승" if fp > 0.3 else "강한 하락" if fp < -1 else "하락" if fp < -0.3 else "보합"
        add("나스닥 선물(24h)", f"{fp:+.2f}%", sig, vote, 1.5)
        fut_vote = vote
    else:
        fut_vote = 0

    # 2) 미국 반도체(SOX) — 간밤 모멘텀(이미 일부 반영)
    cp = (ks.get("change_pct") or {})
    sox = cp.get("SOX")
    if sox is not None:
        vote = _clamp(sox / 3.0 * (0.5 if cp.get("absorbed") else 1.0))  # 소진 시 절반
        add("미국 반도체(SOX)", f"{sox:+.1f}%" + (" (반영됨)" if cp.get("absorbed") else ""),
            "미 반도체 강세" if sox > 1 else "미 반도체 약세" if sox < -1 else "중립", vote, 0.7)

    # 3) VKOSPI 공포 — 변동성·꼬리위험(고=위험회피)
    vv = vk.get("value")
    if vv is not None:
        vote = -1.0 if vv >= 30 else (-0.5 if vv >= 25 else (0.5 if vv < 18 else 0))
        sig = "극단공포·꼬리위험" if vv >= 40 else "공포·위험회피" if vv >= 30 else "보통" if vv >= 20 else "안정"
        add("VKOSPI 변동성", f"{vv} ({vk.get('status','')})", sig, vote, 1.2)

    # 4) 과열/차익실현
    z = oh.get("danger_z")
    if z is not None:
        vote = -2.0 if z >= 2 else (-1.0 if z >= 1 else (1.0 if z <= -2 else 0))
        add("과열/차익실현", f"z {z:+.1f}σ · heat {oh.get('avg_heat','-')}",
            "과열·차익 우선" if z >= 1.5 else "과열 주의" if z >= 0.7 else "침체·저가매수" if z <= -1.5 else "중립", vote, 1.5)

    # 5) 반도체 디커플링 + 외인수급
    reg = (sd.get("regime") or {})
    assets = sd.get("assets") or []
    f5 = sum((a.get("foreign_5d_eok") or 0) for a in assets)
    decoup_high = reg.get("key") == "high"
    if reg or assets:
        vote = -2.0 if f5 < -5000 else (-1.0 if f5 < -500 else (1.0 if f5 > 500 else 0))
        add("반도체 디커플링·외인", f"디커플링 {reg.get('key','-')} · 외인5일 {f5/10000:+.1f}조" if f5 else f"디커플링 {reg.get('key','-')}",
            reg.get("foreign_bias") or ("외인 매도·하방" if f5 < 0 else "외인 매수·상방"), vote, 1.5)

    # 6) 시장감정(KR) — 역발상(탐욕=경계)
    kr = (se.get("kr") or {})
    ss = kr.get("score")
    if ss is not None:
        vote = -1.0 if ss >= 70 else (-0.5 if ss >= 60 else (1.0 if ss <= 30 else (0.5 if ss <= 40 else 0)))
        add("시장감정(KR)", f"{ss} {kr.get('label','')}",
            "극탐욕·추격주의" if ss >= 70 else "탐욕·경계" if ss >= 60 else "극공포·저가매수" if ss <= 30 else "공포" if ss <= 40 else "중립", vote, 0.8)

    # 7) 엔/달러 — 캐리위험
    yen = (ks.get("yen") or {})
    if yen.get("usdjpy") is not None:
        cr = yen.get("carry_risk")
        c1 = yen.get("chg_1d") or 0
        vote = -2.0 if cr else (-0.5 if c1 > 1 else (0.5 if c1 < -1 else 0))
        add("엔/달러", f"{yen.get('usdjpy')} ({c1:+.2f}%)",
            "엔캐리 청산 위험" if cr else "엔약세(경계)" if c1 > 1 else "엔강세" if c1 < -1 else "안정", vote, 0.8)

    # 8) 유동성 스트레스
    lc = lq.get("overall_color")
    if lc:
        vote = 0.5 if lc == "green" else (-1.0 if lc == "amber" else -2.0)
        add("유동성 스트레스", lq.get("overall", ""), "양호" if lc == "green" else "스트레스", vote, 1.0)

    # 9) FX 스왑(외환)
    fxs = fx.get("status")
    if fxs:
        vote = 0 if "🟢" in fxs else (-0.5 if "🟡" in fxs else -2.0)
        add("외환 스왑", fxs, "정상" if "🟢" in fxs else "주의" if "🟡" in fxs else "경색", vote, 0.8)

    # 10) 지정학
    geo = dj.get("geopolitical") or {}
    gr = geo.get("risk_score")
    if gr is not None:
        vote = 0 if gr < 40 else (-0.5 if gr < 60 else -1.5)
        add("지정학 리스크", f"{gr} {geo.get('risk_level','')}", "낮음" if gr < 40 else "관심" if gr < 60 else "고조", vote, 0.7)

    # 11) 통화량 M2 (완화=우호 배경)
    ser = m2.get("series") or []
    if len(ser) >= 13:
        try:
            last = ser[-1].get("value") or ser[-1].get("m2") or ser[-1].get("val")
            yoy = ser[-13].get("value") or ser[-13].get("m2") or ser[-13].get("val")
            g = (float(last) / float(yoy) - 1) * 100 if (last and yoy) else None
            if g is not None:
                vote = 0.5 if g > 3 else (-0.5 if g < 0 else 0)
                add("통화량 M2", f"YoY {g:+.1f}%", "완화·우호" if g > 3 else "긴축" if g < 0 else "중립", vote, 0.5)
        except Exception:
            pass

    # 12) 임박 실적/이벤트 (실적 앞 관망)
    near = [e for e in (ke.get("events") or []) if 0 <= (e.get("dday") or 99) <= 5 and "실적" in (e.get("type") or "")]
    if near:
        e0 = near[0]
        add("임박 실적", f"{e0['name']} {e0['type']} D-{e0['dday']}", "실적 앞 관망·경계", -0.5, 0.5)

    if not factors:
        print("[ERROR] 요소 없음 — 입력 JSON 확인")
        return 1

    # 종합 점수
    wsum = sum(f["weight"] for f in factors)
    score = round(sum(f["vote"] * f["weight"] for f in factors) / wsum, 2) if wsum else 0
    if score >= 0.6:
        bias, bcol = "🟢🟢 강한 상방", "green"
    elif score >= 0.2:
        bias, bcol = "🟢 상방 우위", "green"
    elif score > -0.2:
        bias, bcol = "⚪ 중립·혼조", "gray"
    elif score > -0.6:
        bias, bcol = "🔴 하방 우위", "red"
    else:
        bias, bcol = "🔴🔴 강한 하방", "red"

    # 디커플링 경고: 선물 방향 vs 종합 방향 불일치 또는 디커플링 레짐 high
    fut_dir = 1 if fut_vote > 0.2 else (-1 if fut_vote < -0.2 else 0)
    comp_dir = 1 if score > 0.2 else (-1 if score < -0.2 else 0)
    mismatch = fut_dir != 0 and comp_dir != 0 and fut_dir != comp_dir
    decoup_warn = mismatch or decoup_high
    warn_msg = None
    if decoup_warn:
        warn_msg = ("⚠️ 디커플링 경고 — " +
                    ("선물(개장 전)과 종합 신호 방향 불일치. " if mismatch else "") +
                    ("반도체 디커플링 큼(한국 고유 수급 지배). " if decoup_high else "") +
                    "선물 기반 OHLC 예측 신뢰도 낮음.")

    factors.sort(key=lambda f: abs(f["vote"] * f["weight"]), reverse=True)
    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "base_date": (ks.get("current") or {}).get("kospi_asof") or cp.get("kospi_asof") or now.strftime("%Y-%m-%d"),
        "kospi_now": (ks.get("current") or {}).get("KOSPI"),
        "score": score, "bias": bias, "bias_color": bcol,
        "futures_pct": fp, "futures_bias": "상승" if fut_dir > 0 else "하락" if fut_dir < 0 else "보합",
        "decoupling_warn": decoup_warn, "warn_msg": warn_msg,
        "factors": factors,
        "note": ("개별 지표를 요소별 방향(−2~+2)·가중치로 종합한 다음 거래일 방향 편향. "
                 "선물↔종합 불일치 시 디커플링 경고. 통계 종합이며 투자자문 아님."),
    }
    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"종합 bias: {bias} (score {score:+.2f}) · 선물 {out['futures_bias']} · 경고 {decoup_warn}")
    for f in factors:
        print(f"  {f['name']:16s} {f['signal'][:16]:16s} vote {f['vote']:+.1f} ×{f['weight']}")

    # 텔레그램: 강한 방향 or 디커플링 경고 시(하루 1회 중복방지)
    try:
        from core import send_message, get_secret, load_state, save_state
        if get_secret("TELEGRAM_FINANCE_BOT_TOKEN") and (abs(score) >= 0.6 or decoup_warn):
            st = load_state("composite_signal", default={})
            sig = f"{now:%Y-%m-%d}|{bias}|{decoup_warn}"
            if st.get("sig") != sig:
                body = f"🧭 종합 예측 신호 — {bias} (score {score:+.2f})\n선물: {out['futures_bias']}"
                if warn_msg:
                    body += f"\n{warn_msg}"
                top = factors[:4]
                body += "\n" + "\n".join(f"• {t['name']}: {t['signal']} ({t['vote']:+.1f})" for t in top)
                if send_message(body):
                    save_state("composite_signal", {"sig": sig})
                    print("  📨 종합 신호 경보 발송")
    except Exception as e:
        print(f"  [WARN] 텔레그램 실패: {e}")

    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
