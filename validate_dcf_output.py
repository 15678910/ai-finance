"""
DCF 출력 검증기 — docs/dcf_valuations.json 이 대시보드에 올라가기 전 sanity check
==================================================================================
anthropics/financial-services 의 dcf-model 스킬이 엑셀 모델에 대해 하는 검증
(scripts/validate_dcf.py) 을 우리 JSON 출력에 맞게 옮긴 것.

왜 필요한가 (2026-09-19 사고)
  dcf_valuator.py 가 음수 적정가(-1,597,114원)·-537% 괴리를 그대로 내보냈고 워크플로는
  초록색이었다. 계산이 끝난 뒤 결과를 한 번 더 보는 단계가 없었다. 이 스크립트는
  워크플로에서 `|| echo` 없이 실행돼, 규칙 위반이 있으면 실행을 빨갛게 만든다
  (커밋 스텝은 if: always() 라 데이터는 그대로 올라간다 — 알림만 켜진다).

규칙
  ERROR (exit 1)
    · 파일 파싱 실패 · valuations 비어 있음 · generated_at 이 --max-age-hours 초과
    · 적용 종목: fair_price_dcf ≤ 0 · upside_pct ≤ -100 · wacc ≤ 영구성장률 · 주식수 ≤ 0
      · terminal_share_pct 가 (0,100) 밖 · 숫자 필드에 NaN/inf
    · 부적합 종목: dcf_reason 비어 있음
  WARN (exit 0, 로그만)
    · terminal_share_pct > 75 (터미널 과의존)  · |terminal_method_gap_pct| > 30
    · low_confidence / method_mismatch 종목 수

사용법: python validate_dcf_output.py [경로] [--max-age-hours 36]
"""

import argparse
import json
import math
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PATH = os.path.join(BASE_DIR, "docs", "dcf_valuations.json")

TERMINAL_SHARE_WARN_PCT = 75.0
METHOD_GAP_WARN_PCT = 30.0
NUMERIC_FIELDS = ("fair_price_dcf", "upside_pct", "wacc_pct", "growth_rate_pct", "beta",
                  "terminal_share_pct", "terminal_method_gap_pct", "fair_price_exit",
                  "fair_price_blend", "shares_outstanding", "fcf_base", "current_price")


def _bad_number(v) -> bool:
    return isinstance(v, float) and (math.isnan(v) or math.isinf(v))


def validate(data: dict, now: datetime = None, max_age_hours: float = 36.0) -> tuple:
    """(errors, warnings) — 각각 사람이 읽는 문자열 목록. 순수 함수(테스트 대상)."""
    errors, warnings = [], []
    now = now or datetime.now(KST)

    vals = data.get("valuations")
    if not isinstance(vals, list) or not vals:
        errors.append("valuations 가 비어 있거나 리스트가 아님")
        return errors, warnings

    # 신선도 — 스크립트가 예전 파일을 그대로 두고 끝났는지
    gen = str(data.get("generated_at") or "")
    try:
        gen_dt = datetime.strptime(gen[:19], "%Y-%m-%d %H:%M:%S").replace(tzinfo=KST)
        age_h = (now - gen_dt).total_seconds() / 3600
        if age_h > max_age_hours:
            errors.append(f"generated_at {gen} — {age_h:.0f}시간 경과 (허용 {max_age_hours:.0f}h)")
    except ValueError:
        errors.append(f"generated_at 파싱 불가: {gen!r}")

    tg = float((data.get("assumptions") or {}).get("terminal_growth_pct") or 0)

    n_app = n_inap = n_lowconf = n_mismatch = 0
    for v in vals:
        name = v.get("name") or v.get("ticker") or "?"
        for f in NUMERIC_FIELDS:
            if _bad_number(v.get(f)):
                errors.append(f"{name}: {f} 가 NaN/inf")

        if v.get("dcf_applicable", True) is False:
            n_inap += 1
            if not (v.get("dcf_reason") or "").strip():
                errors.append(f"{name}: 부적합인데 dcf_reason 없음")
            continue

        n_app += 1
        fp = v.get("fair_price_dcf")
        up = v.get("upside_pct")
        if fp is None or fp <= 0:
            errors.append(f"{name}: fair_price_dcf={fp} (양수여야 함 — 음수면 부적합 처리돼야)")
        if up is None or up <= -100:
            errors.append(f"{name}: upside_pct={up} (-100% 이하는 음수 적정가의 흔적)")
        w = v.get("wacc_pct")
        if w is None or w <= tg:
            errors.append(f"{name}: wacc_pct={w} ≤ 영구성장률 {tg} (영구성장법 발산)")
        sh = v.get("shares_outstanding")
        if sh is None or sh <= 0:
            errors.append(f"{name}: shares_outstanding={sh}")
        ts = v.get("terminal_share_pct")
        if ts is not None:
            if not (0 < ts < 100):
                errors.append(f"{name}: terminal_share_pct={ts} 범위 밖")
            elif ts > TERMINAL_SHARE_WARN_PCT:
                warnings.append(f"{name}: 터미널 비중 {ts:.0f}% > {TERMINAL_SHARE_WARN_PCT:.0f}% (터미널 가정 과의존)")
        gap = v.get("terminal_method_gap_pct")
        if gap is not None and abs(gap) > METHOD_GAP_WARN_PCT:
            warnings.append(f"{name}: 영구성장법 vs 출구배수법 괴리 {gap:+.0f}%")
        if v.get("low_confidence"):
            n_lowconf += 1
        if v.get("method_mismatch"):
            n_mismatch += 1

    warnings.insert(0, f"적용 {n_app} · 부적합 {n_inap} · 신뢰도낮음 {n_lowconf} · 방법불일치 {n_mismatch}")
    return errors, warnings


def main() -> int:
    ap = argparse.ArgumentParser(description="DCF 출력 sanity check")
    ap.add_argument("path", nargs="?", default=DEFAULT_PATH)
    ap.add_argument("--max-age-hours", type=float, default=36.0)
    args = ap.parse_args()

    try:
        with open(args.path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[VALIDATE] ERROR 파일 읽기 실패 {args.path}: {e}")
        return 1

    errors, warnings = validate(data, max_age_hours=args.max_age_hours)
    for w in warnings:
        print(f"[VALIDATE] WARN  {w}")
    for e in errors:
        print(f"[VALIDATE] ERROR {e}")
    if errors:
        print(f"[VALIDATE] 실패 — 오류 {len(errors)}건. 위 규칙 위반은 dcf_valuator.py 회귀 신호.")
        return 1
    print(f"[VALIDATE] 통과 — 오류 0건 · 경고 {max(len(warnings) - 1, 0)}건")
    return 0


if __name__ == "__main__":
    sys.exit(main())
