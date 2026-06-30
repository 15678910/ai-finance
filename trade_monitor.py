"""
수출입 동향 자동 수집 — 관세청 무역통계 OpenAPI (data.go.kr)
================================================================
매월 수동 갱신하던 docs/trade_data.json을 관세청 OpenAPI로 자동 갱신.
- 총계(총수출·총수입·무역수지) + 주요 HS 품목별 수출·YoY·월별 추이.
- 발표 주기상 월 1회(+ 월초 며칠) 실행으로 충분.

필요: 공공데이터포털(data.go.kr)에서 '관세청_품목별 국가별 수출입실적(GW)' 활용신청 →
      서비스키를 GitHub Secrets에 CUSTOMS_API_KEY 로 등록(디코딩 키 권장).
엔드포인트: http://apis.data.go.kr/1220000/nitemtrade/getNitemtradeList
출력: docs/trade_data.json  (renderTradeData 스키마 호환)
🚨 정보용·투자자문 아님. 첫 실행 로그로 단위(천달러/달러)·필드명 보정.
"""

import json
import os
import sys
import urllib.request
import urllib.parse
import urllib.error
import xml.etree.ElementTree as ET
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "trade_data.json")
ENDPOINT = "http://apis.data.go.kr/1220000/nitemtrade/getNitemtradeList"

# 주요 HS 4단위 품목 (이름, HS부호, 카테고리) — renderTradeData catColor 키와 일치
HS_ITEMS = [
    ("반도체", "8542", "IT/전자"),
    ("컴퓨터", "8471", "IT/전자"),
    ("무선통신기기", "8517", "IT/전자"),
    ("승용차", "8703", "자동차"),
    ("자동차부품", "8708", "자동차"),
    ("석유제품", "2710", "에너지/화학"),
    ("선박", "8901", "조선"),
]

# 관세청 금액 단위 보정: nitemtrade expDlr이 '천달러'면 억달러=값/1e5, '달러'면 값/1e8.
# 첫 실행 raw 로그를 보고 둘 중 하나로 확정.
SCALE_TO_EOK = 1 / 1e5   # 기본: 천달러 가정


def _key():
    k = os.environ.get("CUSTOMS_API_KEY")
    if k:
        return k.strip()
    try:
        from core import get_secret
        return (get_secret("CUSTOMS_API_KEY") or "").strip() or None
    except Exception:
        return None


def _yymm(dt):
    return dt.strftime("%Y%m")


def _months_back(base, n):
    y, m = base.year, base.month - n
    while m <= 0:
        m += 12
        y -= 1
    return datetime(y, m, 1, tzinfo=KST)


# 서비스키 인코딩 자동 감지: Encoding 키는 raw 그대로, Decoding 키는 quote 필요.
#   어느 쪽을 등록했는지 모르므로 첫 호출에서 둘 다 시도 후 성공 방식을 고정.
_KEY_RAW = {"mode": None}   # None=미정, True=raw 부착, False=quote


def _build_url(api_key, base, raw):
    qs = urllib.parse.urlencode(base)
    sk = api_key if raw else urllib.parse.quote(api_key, safe="")
    return f"{ENDPOINT}?serviceKey={sk}&{qs}"


def _try(url):
    """(items, errmsg). data.go.kr 인증오류 본문도 실패로 처리."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        raw = urllib.request.urlopen(req, timeout=20).read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        body = ""
        try:
            body = e.read().decode("utf-8", "replace")[:200]   # data.go.kr은 401 본문에 사유 명시
        except Exception:
            pass
        return None, f"HTTP {e.code} · {body}"
    except Exception as e:
        return None, str(e)[:80]
    up = raw.upper()
    if any(k in up for k in ("NOT_REGISTERED", "SERVICE_KEY_IS", "LIMITED_NUMBER", "DEADLINE", "ACCESS_DENIED")):
        auth = raw[:140]
        return None, f"인증오류 {auth}"
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return None, f"파싱실패 {raw[:120]}"
    rc = root.findtext(".//resultCode") or root.findtext(".//returnReasonCode")
    if rc not in (None, "00", "0"):
        return None, f"code={rc} {root.findtext('.//resultMsg') or root.findtext('.//returnAuthMsg') or raw[:100]}"
    items = [{c.tag: (c.text or "").strip() for c in it} for it in root.iter("item")]
    return items, None


def fetch(api_key, hs, strt, end):
    """getNitemtradeList 호출 → [item dict]. hs=''면 총계 포함 전체. 키 인코딩 자동 감지."""
    base = {"strtYymm": strt, "endYymm": end, "cntyCd": "", "hsSgn": hs}
    modes = [_KEY_RAW["mode"]] if _KEY_RAW["mode"] is not None else [True, False]
    last = ""
    for raw in modes:
        items, err = _try(_build_url(api_key, base, raw))
        if err is None:
            if _KEY_RAW["mode"] is None:
                _KEY_RAW["mode"] = raw
                print(f"    [키 인코딩] {'raw(Encoding 키)' if raw else 'quote(Decoding 키)'} 사용")
            return items or []
        last = err
    print(f"    [실패] hs={hs}: {last}")
    return []


def _num(d, *keys):
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            try:
                return float(str(v).replace(",", ""))
            except ValueError:
                continue
    return None


def _eok(v):
    return round(v * SCALE_TO_EOK, 2) if v is not None else None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    api_key = _key()
    if not api_key:
        print("[SKIP] CUSTOMS_API_KEY 미설정 — 로컬/미설정 시 스킵(기존 trade_data.json 보존).")
        return 0

    now = datetime.now(KST)
    # 최근 14개월 범위 (최신월 + 전년동월 YoY 확보)
    end = _yymm(now)
    strt = _yymm(_months_back(now, 14))
    print(f"[관세청] {strt}~{end} 조회")

    # ── 1) 총계(전체) — hsSgn 공란 ──
    total_rows = fetch(api_key, "", strt, end)
    # 월별 총계만 추출(statKor에 '총계' 또는 hsCd 공란)
    by_month = {}   # 'YYYYMM' -> {exp, imp, bal}
    for d in total_rows:
        nm = d.get("statKor", "") or d.get("statCd", "")
        hs = d.get("hsCd", "") or d.get("hsSgn", "")
        if "총계" in nm or hs in ("", "00", "0", "총계"):
            ym = (d.get("year") or d.get("yymm") or "").replace("-", "").replace(".", "")[:6]
            if ym:
                by_month[ym] = {
                    "exp": _num(d, "expDlr", "expUsd", "expAmt"),
                    "imp": _num(d, "impDlr", "impUsd", "impAmt"),
                    "bal": _num(d, "balPayments", "balPaymentsDlr", "trBal"),
                }
    if total_rows[:1]:
        print(f"  [raw 총계 샘플] {total_rows[0]}")   # 단위·필드 보정용

    months = sorted(by_month.keys())
    if not months:
        print("[ERROR] 총계 데이터 없음 — 필드명/엔드포인트 보정 필요(위 raw 샘플 참조). 기존 파일 보존.")
        return 1
    latest = months[-1]
    cur = by_month[latest]
    prev_year = by_month.get(str(int(latest) - 100))   # 전년 동월
    exp_eok = _eok(cur["exp"])
    yoy = round((cur["exp"] / prev_year["exp"] - 1) * 100, 1) if (prev_year and prev_year.get("exp")) else None

    # 월별 추이(최근 6개월 총수출 억달러)
    monthly_trend = [{"month": f"{m[:4]}-{m[4:6]}", "total": _eok(by_month[m]["exp"])}
                     for m in months[-6:] if by_month[m].get("exp") is not None]

    # ── 2) 품목별 수출·YoY ──
    products = []
    for name, hs, cat in HS_ITEMS:
        rows = fetch(api_key, hs, strt, end)
        pm = {}
        for d in rows:
            ym = (d.get("year") or d.get("yymm") or "").replace("-", "").replace(".", "")[:6]
            if ym:
                pm[ym] = _num(d, "expDlr", "expUsd", "expAmt")
        cur_v = pm.get(latest)
        py_v = pm.get(str(int(latest) - 100))
        if cur_v is None:
            continue
        p_yoy = round((cur_v / py_v - 1) * 100, 1) if py_v else None
        products.append({"name": name, "export_bn": _eok(cur_v), "yoy_pct": p_yoy,
                         "record": None, "category": cat})
    products.sort(key=lambda p: (p["export_bn"] or 0), reverse=True)

    # 자동 인사이트(숫자 기반 — 분석 날조 금지)
    insights = []
    if products:
        top = products[0]
        insights.append(f"{top['name']} 수출 {top['export_bn']}억달러"
                        + (f" ({top['yoy_pct']:+.1f}% YoY)" if top['yoy_pct'] is not None else "") + " — 최대 품목")
    if exp_eok is not None:
        insights.append(f"{latest[:4]}년 {int(latest[4:6])}월 총수출 {exp_eok}억달러"
                        + (f" ({yoy:+.1f}% YoY)" if yoy is not None else ""))
    if cur.get("bal") is not None:
        insights.append(f"무역수지 {_eok(cur['bal'])}억달러")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "source": "관세청 무역통계 OpenAPI (data.go.kr)",
        "period": f"{latest[:4]}-{latest[4:6]}",
        "period_label": f"{latest[:4]}년 {int(latest[4:6])}월",
        "summary": {
            "total_export_bn": exp_eok,
            "total_export_str": f"{exp_eok:,.1f}억달러" if exp_eok is not None else "—",
            "yoy_pct": yoy, "mom_pct": None, "record": None,
            "note": "관세청 OpenAPI 자동 수집(HS 4단위 기준 — MOTIE 품목분류와 다를 수 있음).",
        },
        "products": products,
        "insights": insights,
        "risk_factors": ["HS 4단위 기준이라 산업부 MTI 품목분류와 수치가 다를 수 있음 — 추세 참고용."],
        "monthly_trend": monthly_trend,
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}  {out['period_label']} 총수출 {exp_eok}억달러 · 품목 {len(products)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
