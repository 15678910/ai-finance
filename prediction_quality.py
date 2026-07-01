"""
예측 평가지수(PQI) + 자동 진단 + 변경 전/후 기준표
========================================================
예측 품질을 단일 지수(PQI 0~100)로 계량하고, 지수가 낮거나 떨어질 때
'무엇을 왜 보완해야 하는지'를 규칙기반으로 자동 진단·권고한다.
모델 로직을 바꿀 때마다 MODEL_VERSION을 올리면 버전별 PQI가 기록되어
'변경 전 vs 변경 후' 개선 여부를 객관적으로 비교할 수 있다.

PQI = 0.45·방향적중 + 0.35·크기점수(MAE역) + 0.20·범위적중  (모델별 → 종합 평균)
보완 권고: auto_safe=True(밴드 σ 확대 등 기계적)만 자동적용 후보, 나머지는 검토 후 적용.

출력: docs/prediction_quality.json
🚨 통계 자기평가. 투자자문 아님.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE_DIR, "docs")
OUTPUT_FILE = os.path.join(DOCS, "prediction_quality.json")

# 모델 로직을 바꿀 때마다 올린다 → 버전별 PQI가 기준표에 기록됨(변경 전/후 비교).
MODEL_VERSION = "v1.1"
VERSION_NOTE = "밴드 확대 — OHLC 예측 밴드 σ 1.0→1.5(SK·삼성). 범위적중·급오차 개선 목표"


def _load(name):
    try:
        with open(os.path.join(DOCS, f"{name}.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _pqi(dir_hit, mae, range_hit):
    """방향·크기·범위 → 0~100. range 없으면 방향·크기만 재정규화."""
    dsc = dir_hit if dir_hit is not None else 50
    msc = max(0, min(100, 110 - (mae if mae is not None else 9) * 14))
    if range_hit is None:
        return round(0.56 * dsc + 0.44 * msc, 1)
    return round(0.45 * dsc + 0.35 * msc + 0.20 * range_hit, 1)


def _grade(pqi):
    if pqi >= 70:
        return "🟢 우수", "green"
    if pqi >= 58:
        return "🟢 양호", "green"
    if pqi >= 48:
        return "🟡 보통", "amber"
    if pqi >= 38:
        return "🟠 미흡", "orange"
    return "🔴 취약", "red"


def _windowed_pqi(entries, n):
    """엔트리 최근 n개로 방향·MAE·범위 재계산 → PQI (추세용)."""
    e = [x for x in entries if x.get("abs_err") is not None][-n:]
    if not e:
        return None
    dir_hit = round(sum(1 for x in e if x.get("dir_ok")) / len(e) * 100)
    mae = round(sum(abs(x["abs_err"]) for x in e) / len(e), 2)
    rh = [x for x in e if "range_hit" in x]
    range_hit = round(sum(1 for x in rh if x.get("range_hit")) / len(rh) * 100) if rh else None
    return _pqi(dir_hit, mae, range_hit)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    pl = _load("prediction_log")
    sk = _load("sk_hynix_forecast")
    sm = _load("samsung_forecast")

    # 엔트리 구조가 달라 (KOSPI는 abs_err/dir이 fut 하위) 균일화한다.
    def _norm_kospi(ents):
        out = []
        for e in ents:
            fu = e.get("fut") or {}
            out.append({"abs_err": fu.get("abs_err"), "dir_ok": fu.get("dir"), "range_hit": None})
        return out

    models = []
    fa = (pl.get("accuracy") or {}).get("fut") or {}
    if fa.get("n"):
        models.append({"name": "KOSPI(선물)", "dir": fa.get("dir_hit_rate"), "mae": fa.get("mae_pct"),
                       "range": None, "n": fa.get("n"), "entries": _norm_kospi(pl.get("entries") or [])})
    for nm, d in (("SK하이닉스", sk), ("삼성전자", sm)):
        a = d.get("accuracy") or {}
        if a.get("n"):
            models.append({"name": nm, "dir": a.get("dir_hit_rate"), "mae": a.get("close_mae_pct"),
                           "range": a.get("range_hit_rate"), "n": a.get("n"), "entries": d.get("entries") or []})
    if not models:
        print("[ERROR] 정확도 데이터 없음")
        return 1

    for m in models:
        m["pqi"] = _pqi(m["dir"], m["mae"], m["range"])
        ent = m.pop("entries")
        m["pqi_recent"] = _windowed_pqi(ent, 7)
        m["pqi_prev"] = _windowed_pqi(ent[:-7], 7) if len(ent) > 7 else None
        valid = [e for e in ent if e.get("abs_err") is not None]
        big = [e for e in valid if abs(e["abs_err"]) > 3]
        m["big_miss_frac"] = round(len(big) / len(valid) * 100) if valid else 0

    pqi_overall = round(sum(m["pqi"] for m in models) / len(models), 1)
    grade, gcol = _grade(pqi_overall)
    # 추세(최근7 vs 이전7 평균)
    rec_vals = [m["pqi_recent"] for m in models if m["pqi_recent"] is not None]
    prev_vals = [m["pqi_prev"] for m in models if m["pqi_prev"] is not None]
    trend = round(sum(rec_vals) / len(rec_vals) - sum(prev_vals) / len(prev_vals), 1) if (rec_vals and prev_vals) else None

    # ── 자동 진단·보완 권고 ──
    dir_avg = sum(m["dir"] for m in models if m["dir"] is not None) / len([m for m in models if m["dir"] is not None])
    mae_avg = sum(m["mae"] for m in models if m["mae"] is not None) / len([m for m in models if m["mae"] is not None])
    range_vals = [m["range"] for m in models if m["range"] is not None]
    range_avg = sum(range_vals) / len(range_vals) if range_vals else None
    big_avg = sum(m["big_miss_frac"] for m in models) / len(models)

    recs = []

    def rec(prio, issue, cause, fix, auto_safe):
        recs.append({"priority": prio, "issue": issue, "cause": cause, "fix": fix, "auto_safe": auto_safe})

    if range_avg is not None and range_avg < 40:
        rec(1, f"범위 적중 낮음 ({range_avg:.0f}%)", "예측 OHLC 밴드가 현재 변동성 레짐 대비 협소",
            "밴드 σ 배수 확대 (예: 1.0→1.4)", True)
    if dir_avg < 55:
        rec(1, f"방향 적중 낮음 ({dir_avg:.0f}% · 동전 수준)", "선물 편중 → 디커플링·차익실현 날 방향 오류",
            "디커플링 경고일은 종합신호(composite)로 방향 보정", False)
    if big_avg > 40:
        rec(1, f"급오차일(±3%↑) 빈발 ({big_avg:.0f}%)", "차익실현·외인수급 급변을 선물신호가 못 봄",
            "composite 디커플링 경고일 예측 신뢰도 하향 표시 + 밴드 추가확대", True)
    if mae_avg > 4.5:
        rec(2, f"종가 오차 큼 (MAE {mae_avg:.1f}%)", "고변동성 레짐 + 20일 베타가 급변 미추종",
            "베타 추정창 단축(20→10) 또는 변동성 가중 회귀", False)
    if trend is not None and trend < -4:
        rec(1, f"PQI 하락 추세 ({trend:+.1f})", "최근 레짐 변화(변동성 급등)로 성능 저하",
            "밴드·베타 재추정 주기 단축 + 디커플링 보정 우선 적용", False)
    if not recs:
        rec(3, "임계 이상 양호", "주요 지표 기준 충족", "현 설정 유지·모니터링", True)
    recs.sort(key=lambda r: r["priority"])

    # ── 변경 전/후 버전 기준표 ──
    versions = []
    try:
        from core import load_state, save_state
        st = load_state("prediction_quality", default={}) or {}
        versions = st.get("versions", [])
        today = now.strftime("%Y-%m-%d")
        if versions and versions[-1]["version"] == MODEL_VERSION:
            v = versions[-1]
            v.setdefault("samples", []).append(pqi_overall)
            v["samples"] = v["samples"][-90:]
            v["avg_pqi"] = round(sum(v["samples"]) / len(v["samples"]), 1)
            v["last_date"] = today
        else:
            versions.append({"version": MODEL_VERSION, "note": VERSION_NOTE,
                             "start_date": today, "last_date": today,
                             "samples": [pqi_overall], "avg_pqi": pqi_overall})
        save_state("prediction_quality", {"versions": versions[-12:]})
    except Exception as e:
        print(f"  [WARN] 버전 기준표 저장 실패: {e}")

    ver_table = [{"version": v["version"], "note": v.get("note", ""), "avg_pqi": v["avg_pqi"],
                  "start_date": v["start_date"], "last_date": v.get("last_date", ""),
                  "n_days": len(v.get("samples", []))} for v in versions]
    delta_vs_prev = None
    if len(ver_table) >= 2:
        delta_vs_prev = round(ver_table[-1]["avg_pqi"] - ver_table[-2]["avg_pqi"], 1)

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "model_version": MODEL_VERSION,
        "pqi_overall": pqi_overall, "grade": grade, "grade_color": gcol, "trend": trend,
        "models": models,
        "diagnosis": {"dir_avg": round(dir_avg, 1), "mae_avg": round(mae_avg, 2),
                      "range_avg": round(range_avg, 1) if range_avg is not None else None,
                      "big_miss_avg": round(big_avg, 1)},
        "recommendations": recs,
        "version_table": ver_table, "delta_vs_prev_version": delta_vs_prev,
        "note": ("PQI=0.45·방향+0.35·크기(MAE역)+0.20·범위 (모델 평균). 지수 하락 시 규칙기반 자동 진단·권고. "
                 "auto_safe=밴드확대 등 기계적 보정만 자동 후보, 모델 로직 변경은 검토 후 적용. 통계 자기평가."),
    }
    os.makedirs(DOCS, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"PQI 종합: {pqi_overall} {grade} (추세 {trend}) · 버전 {MODEL_VERSION}")
    for m in models:
        print(f"  {m['name']:12s} PQI {m['pqi']} (dir {m['dir']} mae {m['mae']} range {m['range']} 급오차 {m['big_miss_frac']}%)")
    print("  권고:")
    for r in recs:
        print(f"   [P{r['priority']}{'·자동가능' if r['auto_safe'] else '·검토'}] {r['issue']} → {r['fix']}")

    # 텔레그램: PQI 취약(<38) 또는 급락(추세<-6) 시 (중복방지)
    try:
        from core import send_message, get_secret, load_state, save_state
        if get_secret("TELEGRAM_FINANCE_BOT_TOKEN") and (pqi_overall < 38 or (trend is not None and trend < -6)):
            st = load_state("pqi_alert", default={})
            sig = f"{now:%Y-%m-%d}|{grade}"
            if st.get("sig") != sig:
                top = recs[0]
                body = (f"📉 예측 평가지수 경보 — PQI {pqi_overall} {grade} (추세 {trend})\n"
                        f"최우선 보완: {top['issue']}\n→ {top['fix']}")
                if send_message(body):
                    save_state("pqi_alert", {"sig": sig})
                    print("  📨 PQI 경보 발송")
    except Exception as e:
        print(f"  [WARN] 텔레그램 실패: {e}")

    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
