"""
재무건전성 지표 — DART 재무제표 기반 (유동비율·부채비율·이자보상배율·EPS 추세)
================================================================================
KRX 시장지표(PER·PBR·EPS)만으로는 볼 수 없는 '재무제표 기반' 안정성을 DART OpenAPI로 수집.

  · 유동비율   = 유동자산 ÷ 유동부채 × 100    (단기 지급능력, 150%↑ 양호)
  · 부채비율   = 부채총계 ÷ 자본총계 × 100    (재무 안정성, 100%↓ 양호)
  · 이자보상배율 = 영업이익 ÷ 이자비용        (1 미만 = 이자도 못 갚는 상태)
  · EPS 추세   = 최근 3~4개 사업연도 EPS + CAGR + 연속증가 여부 (성장의 '일관성')

DART는 IP 차단이 없어 GitHub Actions에서 실행 가능(로컬은 DART_API_KEY 없으면 스킵).
연 1회 갱신되는 사업보고서 기반이라 자주 안 바뀜 — 하루 1회 실행으로 충분.

출력: docs/dart_financials.json
🚨 정보 제공용 · 투자자문 아님. 연결(CFS) 우선, 없으면 별도(OFS).
"""

import io
import json
import os
import re
import sys
import time
import urllib.request
import zipfile
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "dart_financials.json")
DART = "https://opendart.fss.or.kr/api"
UA = {"User-Agent": "Mozilla/5.0 (compatible; ai-finance-dashboard)"}

WATCH = [
    ("000660", "SK하이닉스"), ("005930", "삼성전자"), ("108490", "로보티즈"),
    ("003550", "LG"), ("066570", "LG전자"), ("042700", "한미반도체"), ("009150", "삼성전기"),
    ("373220", "LG에너지솔루션"),
]
N_YEARS = 4                                   # 최근 4개 사업연도 시도

# IFRS 표준 account_id (우선) / 계정명 키워드 (폴백)
ACCOUNTS = {
    "current_assets":  (["ifrs-full_CurrentAssets"], ["유동자산"]),
    "current_liab":    (["ifrs-full_CurrentLiabilities"], ["유동부채"]),
    "liabilities":     (["ifrs-full_Liabilities"], ["부채총계"]),
    "equity":          (["ifrs-full_Equity"], ["자본총계"]),
    "op_income":       (["dart_OperatingIncomeLoss", "ifrs-full_ProfitLossFromOperatingActivities"], ["영업이익"]),
    "net_income":      (["ifrs-full_ProfitLoss"], ["당기순이익"]),
    "interest_exp":    (["ifrs-full_InterestExpense"], ["이자비용", "금융원가", "금융비용"]),
    "eps":             (["ifrs-full_BasicEarningsLossPerShare"], ["기본주당이익", "주당이익", "기본주당순이익"]),
}


def _key():
    k = os.environ.get("DART_API_KEY")
    if k:
        return k.strip()
    try:
        from core import get_secret
        return (get_secret("DART_API_KEY") or "").strip() or None
    except Exception:
        return None


def _get(url, timeout=25):
    return urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=timeout).read()


def corp_code_map(key):
    """DART corpCode.zip → {종목코드6: 고유번호8}"""
    raw = _get(f"{DART}/corpCode.xml?crtfc_key={key}", timeout=60)
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        xml = z.read(z.namelist()[0]).decode("utf-8", "replace")
    m = {}
    for blk in re.findall(r"<list>(.*?)</list>", xml, re.S):
        sc = re.search(r"<stock_code>\s*(\S+)\s*</stock_code>", blk)
        cc = re.search(r"<corp_code>\s*(\S+)\s*</corp_code>", blk)
        if sc and cc and sc.group(1) and sc.group(1) != " ":
            m[sc.group(1).strip()] = cc.group(1).strip()
    return m


def _num(s):
    try:
        v = str(s).replace(",", "").strip()
        if v in ("", "-"):
            return None
        return float(v)
    except Exception:
        return None


def fetch_year(key, corp, year):
    """사업보고서(11011) 전체 계정 → 지표 dict. 연결(CFS) 우선, 실패 시 별도(OFS)."""
    for fs in ("CFS", "OFS"):
        try:
            url = (f"{DART}/fnlttSinglAcntAll.json?crtfc_key={key}&corp_code={corp}"
                   f"&bsns_year={year}&reprt_code=11011&fs_div={fs}")
            d = json.loads(_get(url).decode("utf-8", "replace"))
        except Exception:
            continue
        if d.get("status") != "000" or not d.get("list"):
            continue
        out = {"fs_div": fs}
        for field, (ids, names) in ACCOUNTS.items():
            val = None
            for it in d["list"]:                                    # 1순위: 표준 account_id
                if it.get("account_id") in ids:
                    val = _num(it.get("thstrm_amount"))
                    if val is not None:
                        break
            if val is None:                                         # 2순위: 계정명 키워드
                for it in d["list"]:
                    nm = re.sub(r"[\s\.\dIVXⅠ-Ⅹ]", "", str(it.get("account_nm", "")))
                    if any(nm == n or nm.startswith(n) for n in names):
                        val = _num(it.get("thstrm_amount"))
                        if val is not None:
                            break
            out[field] = val
        if out.get("equity") or out.get("current_assets"):
            return out
    return None


def ratios(y):
    """원시 계정 → 재무비율."""
    ca, cl = y.get("current_assets"), y.get("current_liab")
    li, eq = y.get("liabilities"), y.get("equity")
    op, ie = y.get("op_income"), y.get("interest_exp")
    r = {}
    r["current_ratio"] = round(ca / cl * 100, 1) if (ca and cl and cl > 0) else None
    r["debt_ratio"] = round(li / eq * 100, 1) if (li and eq and eq > 0) else None
    if op is not None and ie and ie > 0:
        r["interest_coverage"] = round(op / ie, 1)
    elif op is not None and (ie in (0, None)):
        r["interest_coverage"] = None                               # 이자비용 미공시/0 → 판정 보류
    else:
        r["interest_coverage"] = None
    return r


def health(latest):
    """3지표 신호등 — 각 양호2/보통1/주의0, 합계로 등급."""
    cr, dr, ic = latest.get("current_ratio"), latest.get("debt_ratio"), latest.get("interest_coverage")
    sc, flags = 0, []
    n = 0
    if cr is not None:
        n += 1
        sc += 2 if cr >= 150 else 1 if cr >= 100 else 0
        if cr < 100:
            flags.append(f"유동비율 {cr}% (100% 미만 — 단기 지급능력 주의)")
    if dr is not None:
        n += 1
        sc += 2 if dr <= 100 else 1 if dr <= 200 else 0
        if dr > 200:
            flags.append(f"부채비율 {dr}% (200% 초과 — 재무 부담 큼)")
    if ic is not None:
        n += 1
        sc += 2 if ic >= 3 else 1 if ic >= 1 else 0
        if ic < 1:
            flags.append(f"이자보상배율 {ic}배 (1 미만 — 영업이익으로 이자 미충당)")
    if n == 0:
        return {"score": None, "grade": "판정 불가", "color": "gray", "flags": ["재무 항목 미확보"]}
    pct = sc / (n * 2) * 100
    if pct >= 80:
        g, c = "🟢 양호", "green"
    elif pct >= 50:
        g, c = "🟡 보통", "yellow"
    else:
        g, c = "🔴 주의", "red"
    return {"score": round(pct), "grade": g, "color": c, "flags": flags}


def eps_trend(years):
    """EPS 다년 추세: 값·CAGR·연속증가 여부."""
    seq = [(y["year"], y.get("eps")) for y in years if y.get("eps") is not None]
    seq.sort()
    vals = [v for _, v in seq]
    if len(vals) < 2:
        return {"values": [{"year": y, "eps": round(v)} for y, v in seq], "cagr": None,
                "consistent": None, "note": "표본 부족"}
    cagr = None
    if vals[0] > 0 and vals[-1] > 0:
        n = len(vals) - 1
        cagr = round(((vals[-1] / vals[0]) ** (1 / n) - 1) * 100, 1)
    consistent = all(vals[i] > vals[i - 1] for i in range(1, len(vals)))
    return {"values": [{"year": y, "eps": round(v)} for y, v in seq], "cagr": cagr,
            "consistent": consistent,
            "note": ("매년 증가 (일관된 성장)" if consistent else "증감 혼재 — 일시적 반등 가능성 확인 필요")}


def screener_health(key, cmap, now):
    """퀀트 스크리너 통과 후보에 재무건전성 부착.
    전 종목 DART 조회는 수천 콜이라 비현실적 → 이미 KRX로 걸러진 20여 종목만 조회(최신 1개 연도)."""
    path = os.path.join(BASE_DIR, "docs", "stock_screener.json")
    try:
        with open(path, encoding="utf-8") as f:
            sc = json.load(f)
    except Exception:
        print("  [INFO] stock_screener.json 없음 — 스크리너 건전성 스킵")
        return {}, None
    out, cands = {}, (sc.get("stocks") or [])
    for s in cands:
        code = s.get("code")
        corp = cmap.get(code)
        if not corp:
            continue
        for y in (now.year - 1, now.year - 2):          # 최신 사업연도 우선, 없으면 직전
            r = fetch_year(key, corp, y)
            time.sleep(0.25)
            if r:
                rr = ratios(r)
                out[code] = {"fy": y, "fs_div": r.get("fs_div"), **rr, "health": health(rr)}
                break
    ok = sum(1 for v in out.values() if (v.get("interest_coverage") or 0) >= 3 and (v.get("debt_ratio") or 999) <= 150)
    print(f"  스크리너 후보 {len(cands)}종목 중 {len(out)}종목 재무 확보 · 건전성 통과(이자보상≥3·부채≤150%) {ok}종목")
    return out, sc.get("asof")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    key = _key()
    if not key:
        print("[SKIP] DART_API_KEY 미설정 — 로컬 스킵(워크플로에서 실행). 기존 JSON 보존.")
        return 0

    now = datetime.now(KST)
    try:
        cmap = corp_code_map(key)
        print(f"  DART 고유번호 매핑 {len(cmap):,}건 로드")
    except Exception as e:
        print(f"[ERROR] corpCode 로드 실패: {e}")
        return 1

    yrs = [now.year - k for k in range(1, N_YEARS + 1)]              # 최신 사업연도부터 역순
    stocks = []
    for code, name in WATCH:
        corp = cmap.get(code)
        if not corp:
            print(f"  [WARN] {name}({code}) DART 고유번호 없음")
            continue
        years = []
        for y in yrs:
            r = fetch_year(key, corp, y)
            time.sleep(0.25)                                        # DART 예의(호출 간격)
            if not r:
                continue
            rec = {"year": y, **r, **ratios(r)}
            years.append(rec)
        if not years:
            print(f"  [WARN] {name} 재무제표 없음")
            continue
        years.sort(key=lambda x: -x["year"])
        latest = years[0]
        st = {
            "code": code, "name": name, "corp_code": corp,
            "fs_div": latest.get("fs_div"), "fy": latest["year"],
            "current_ratio": latest.get("current_ratio"),
            "debt_ratio": latest.get("debt_ratio"),
            "interest_coverage": latest.get("interest_coverage"),
            "op_income": latest.get("op_income"), "equity": latest.get("equity"),
            "health": health(latest),
            "eps_trend": eps_trend(years),
            "history": [{"year": y["year"], "current_ratio": y.get("current_ratio"),
                         "debt_ratio": y.get("debt_ratio"),
                         "interest_coverage": y.get("interest_coverage")} for y in years],
        }
        stocks.append(st)
        et = st["eps_trend"]
        print(f"  {name}: FY{latest['year']}({latest.get('fs_div')}) 유동 {st['current_ratio']}% "
              f"부채 {st['debt_ratio']}% 이자보상 {st['interest_coverage']}배 → {st['health']['grade']} "
              f"| EPS CAGR {et.get('cagr')}% 연속증가 {et.get('consistent')}")

    if not stocks:
        print("[ERROR] 전 종목 실패 — 기존 파일 보존.")
        return 1

    sh, sh_asof = screener_health(key, cmap, now)      # 스크리너 후보 건전성

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "stocks": stocks,
        "screener_health": sh, "screener_asof": sh_asof,
        "criteria": {"유동비율": "≥150% 양호 · <100% 주의", "부채비율": "≤100% 양호 · >200% 주의",
                     "이자보상배율": "≥3배 양호 · <1배 경고(영업이익<이자)"},
        "note": ("DART 사업보고서(연간, 연결 우선) 기반 재무건전성. "
                 "유동비율=유동자산/유동부채, 부채비율=부채총계/자본총계, 이자보상배율=영업이익/이자비용. "
                 "이자비용 미공시 종목은 배율 '—'(판정 보류). EPS는 사업보고서 기본주당이익. "
                 "연 1회 갱신되는 후행 지표 — 최근 분기 상황은 반영 안 됨. 정보 제공용 · 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE} ({len(stocks)}종목)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
