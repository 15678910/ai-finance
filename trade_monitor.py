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

# 주요 품목 (이름, [HS 4단위 코드들], 카테고리) — 산업부 MTI 분류 근사 위해 관련 HS 합산.
#   예) 반도체=집적회로(8542)+개별소자(8541), 컴퓨터=본체(8471)+SSD(8523, 한국 수출 대부분).
HS_ITEMS = [
    ("반도체", ["8541", "8542"], "IT/전자"),
    ("컴퓨터·SSD", ["8471", "8523"], "IT/전자"),
    ("무선통신기기", ["8517"], "IT/전자"),
    ("승용차", ["8703"], "자동차"),
    ("자동차부품", ["8708"], "자동차"),
    ("석유제품", ["2710"], "에너지/화학"),
    ("선박", ["8901", "8904", "8905"], "조선"),
]
HS_ITEMS_MAP = {name: "+".join(codes) for name, codes, _ in HS_ITEMS}

# 관세청 nitemtrade expDlr은 '달러' 단위(raw 샘플 확인: 185527777403=$185.5B) → 억달러=값/1e8.
SCALE_TO_EOK = 1 / 1e8


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
    # 정규화: Encoding 키든 Decoding 키든 unquote→quote로 올바른 인코딩 형태 통일.
    sk = api_key if raw else urllib.parse.quote(urllib.parse.unquote(api_key), safe="")
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
    # 정규화(False) 먼저, 안 되면 raw(True) 폴백
    modes = [_KEY_RAW["mode"]] if _KEY_RAW["mode"] is not None else [False, True]
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
    end = _yymm(now)
    # API 제약: 조회기간 1년 이내 → 두 창으로 분할(최근 12개월 + 전년동월 YoY 기준).
    w1s = _yymm(_months_back(now, 11))                                  # 최근 12개월
    w2s, w2e = _yymm(_months_back(now, 14)), _yymm(_months_back(now, 12))  # 전년동월 부근(최신월 -12 커버)
    print(f"[관세청] {w1s}~{end} + {w2s}~{w2e} 조회 (품목별)")

    def _month(d):
        v = (d.get("year") or d.get("yymm") or d.get("baseYymm") or "").replace("-", "").replace(".", "")[:6]
        return v if (v.isdigit() and len(v) == 6) else ""   # year='총계'(전기간 합계) 행 제외

    def _is_total_row(d):
        return any("총계" in str(v) for v in d.values())

    def _by_month(rows):
        """HS 한 품목의 월별 {exp,imp,bal} — 총계행 우선, 없으면 국가별 합산."""
        totals, sums = {}, {}
        for d in rows:
            ym = _month(d)
            if not ym:
                continue
            e, i, b = (_num(d, "expDlr", "expUsd", "expAmt", "expDlrAmt"),
                       _num(d, "impDlr", "impUsd", "impAmt", "impDlrAmt"),
                       _num(d, "balPayments", "balPaymentsDlr", "trBal"))
            if _is_total_row(d):
                totals[ym] = {"exp": e, "imp": i, "bal": b}
            else:
                s = sums.setdefault(ym, {"exp": 0.0, "imp": 0.0, "bal": 0.0})
                s["exp"] += e or 0; s["imp"] += i or 0
                s["bal"] = (s["exp"] - s["imp"])
        return totals or sums

    def _merge(a, b):
        for ym, v in b.items():
            t = a.setdefault(ym, {"exp": 0.0, "imp": 0.0, "bal": 0.0})
            t["exp"] += v.get("exp") or 0
            t["imp"] += v.get("imp") or 0
            t["bal"] = t["exp"] - t["imp"]
        return a

    products, prod_series, raw_logged = [], {}, False
    for name, codes, cat in HS_ITEMS:
        bm = {}
        for hs in codes:                                  # 품목 = 여러 HS 합산
            rows = fetch(api_key, hs, w1s, end) + fetch(api_key, hs, w2s, w2e)
            if rows and not raw_logged:
                print(f"  [raw 샘플 {name}({hs})] rows={len(rows)} 첫행={rows[0]}")
                raw_logged = True
            _merge(bm, _by_month(rows))
        prod_series[name] = bm
        latest_m = max(bm) if bm else None
        if not latest_m:
            continue
        cur_v = bm[latest_m]["exp"]
        py_v = bm.get(str(int(latest_m) - 100), {}).get("exp")
        p_yoy = round((cur_v / py_v - 1) * 100, 1) if (cur_v and py_v) else None
        products.append({"name": name, "export_bn": _eok(cur_v), "yoy_pct": p_yoy,
                         "record": None, "category": cat, "_m": latest_m})

    if not products:
        print("[ERROR] 품목 데이터 없음 — 위 raw 샘플로 필드명 보정 필요. 기존 파일 보존.")
        return 1

    latest = max(p["_m"] for p in products)
    products.sort(key=lambda p: (p["export_bn"] or 0), reverse=True)
    for p in products:
        p.pop("_m", None)

    # 헤드라인 = 반도체(HS8542) 우선, 없으면 최대 품목
    head = next((p for p in products if p["name"] == "반도체"), products[0])
    head_bm = prod_series.get(head["name"], {})
    sorted_m = sorted(head_bm.keys())
    monthly_trend = [{"month": f"{m[:4]}-{m[4:6]}", "total": _eok(head_bm[m]["exp"])}
                     for m in sorted_m[-6:] if head_bm[m].get("exp") is not None]

    insights = [f"{p['name']} 수출 {p['export_bn']}억달러"
                + (f" ({p['yoy_pct']:+.1f}% YoY)" if p['yoy_pct'] is not None else "")
                for p in products[:3]]

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "source": "관세청 무역통계 OpenAPI (data.go.kr)",
        "period": f"{latest[:4]}-{latest[4:6]}",
        "period_label": f"{latest[:4]}년 {int(latest[4:6])}월",
        "summary": {
            "headline_label": f"{head['name']} 수출 (HS {dict(HS_ITEMS_MAP).get(head['name'],'')})",
            "total_export_bn": head["export_bn"],
            "total_export_str": f"{head['export_bn']:,.1f}억달러" if head["export_bn"] is not None else "—",
            "yoy_pct": head["yoy_pct"], "mom_pct": None, "record": None,
            "note": "관세청 OpenAPI 자동 수집(HS 4단위 기준 — 산업부 MTI 품목분류·국가총계와 다를 수 있음).",
        },
        "products": products,
        "insights": insights,
        "risk_factors": ["HS 4단위 기준 — 산업부 MTI 품목분류와 수치 상이. 국가 총수출 총계는 별도 API 필요(추세 참고용)."],
        "monthly_trend": monthly_trend,
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}  {out['period_label']} · {head['name']} {head['export_bn']}억달러 · 품목 {len(products)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
