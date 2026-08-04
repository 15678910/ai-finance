"""통화량·물가·자산 전달경로 모니터

무엇을 보나
----------
"돈을 풀면 물가가 오르고, 물가가 오르면 명목 세수가 늘고, 자산가격도 오른다"는
통념을 **실제 데이터로 측정**한다. 결론부터: 이 관계는 상시 성립하지 않고 시대에 따라
뒤집힌다(아래 VALIDATION 참조). 그래서 단정 대신 '지금 어느 국면인가'를 보여준다.

4개 블록
--------
1) 주요 5개 경제권 통화량(M3) — 미국·유로존·일본·한국·중국. OECD SDMX(키 불필요).
   증가율은 자국통화 기준이라 환율과 무관. 비중은 최신 환율로 달러 환산한 근사치.
2) 전달경로 — 미국 M2→CPI→연방세수의 시차 상관. 구간을 나눠 관계가 언제 성립했는지 표시.
3) 실질주택가격 — BIS 4개국(미·한·일·중). 명목가격÷물가라 '물가를 이겼는가'를 본다.
4) 재정 — 관세수입과 연방 총세수 대비 비중.

VALIDATION (2026-08-04 실측)
  · 미국 M2→CPI: 전체 18개월 시차 r=+0.504 —— 그러나
      1996~2007 r=+0.107 / 2008~2019(QE) r=-0.294 / 2020~2026 r=+0.844
    QE 시기엔 통화량이 폭증해도 물가가 안 올랐다(오히려 역상관). 조건부 관계다.
  · 미국 M2→세수: 12개월 r=+0.251(약함) < CPI→세수 동행 r=+0.485
    → 통화량은 세수에 '직접' 닿지 않고 물가를 매개로 전달된다.
  · 미국 M2→실질주택가격: 한국이 6개월 시차 r=+0.488로 4개국 중 가장 민감.

출력: docs/money_macro.json
🚨 상관관계는 인과가 아니다. 공개 데이터의 규칙기반 요약 · 투자자문 아님.
"""

import csv
import io
import json
import os
import subprocess
import sys
import urllib.parse
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "money_macro.json")

FRED_CSV = "https://fred.stlouisfed.org/graph/fredgraph.csv"
OECD_MONAGG = ("https://sdmx.oecd.org/public/rest/data/OECD.SDD.STES,DSD_STES@DF_MONAGG,"
               "/all?startPeriod={start}&format=csvfilewithlabels")

# OECD 통화총량 대상 (M3=광의통화). XDC=자국통화 단위
AREAS = [("USA", "미국", "🇺🇸", "USD"), ("EA20", "유로존", "🇪🇺", "EUR"),
         ("JPN", "일본", "🇯🇵", "JPY"), ("KOR", "한국", "🇰🇷", "KRW"),
         ("CHN", "중국", "🇨🇳", "CNY")]
# 달러 환산용 야후 티커 (USD당 자국통화. EUR만 역방향이라 별도 처리)
FX_TICKER = {"KRW": "KRW=X", "JPY": "JPY=X", "CNY": "CNY=X", "EUR": "EURUSD=X"}

PROPERTY = [("QUSR628BIS", "미국", "🇺🇸"), ("QKRR628BIS", "한국", "🇰🇷"),
            ("QJPR628BIS", "일본", "🇯🇵"), ("QCNR628BIS", "중국", "🇨🇳")]

# ── ECOS(한국은행) 자동 탐색 설정 ──────────────────────────────────
# 통계표코드를 상수로 박지 않고 '이름 키워드'로 찾는다.
# 한국은행이 코드를 바꿔도 깨지지 않게 하려는 것(m2_monitor.py와 같은 방식).
ECOS_BASE = "https://ecos.bok.or.kr/api"
ECOS_TARGETS = {
    # 4.4 부동산 가격지수 → 서울 아파트 매매
    "seoul_apt": {
        "table_kw": ["주택매매가격", "부동산 가격", "부동산가격"],
        "item_kw_all": ["서울"],              # 항목명에 반드시 포함
        "item_kw_any": ["아파트"],            # 이 중 하나 포함
        "label": "서울 아파트 매매가격지수",
    },
    # 4.2 소비자물가지수 → 총지수 (실질화 분모)
    "kr_cpi": {
        "table_kw": ["소비자물가지수"],
        "item_kw_all": [],
        "item_kw_any": ["총지수", "총 지수"],
        "label": "한국 소비자물가지수",
    },
}


# ── 수집 ────────────────────────────────────────────────────────────
_USE_CURL = None      # None=미판별 / True=urllib 막힌 환경 / False=urllib 사용


def _probe_urllib():
    """urllib이 이 환경에서 외부에 닿는지 '한 번만' 짧게 확인.
    매 요청마다 긴 타임아웃을 기다리면 전체 수집이 수 분씩 지연되므로 결과를 캐시한다."""
    global _USE_CURL
    if _USE_CURL is not None:
        return _USE_CURL
    try:
        req = urllib.request.Request(f"{FRED_CSV}?id=DGS10&cosd=2026-07-01",
                                     headers={"User-Agent": "Mozilla/5.0"})
        urllib.request.urlopen(req, timeout=8).read(64)
        _USE_CURL = False
    except Exception:
        print("  [INFO] urllib 외부 접속 불가 — curl 경로 사용")
        _USE_CURL = True
    return _USE_CURL


def _http(url, timeout=60):
    """외부 CSV 수집. 환경에 따라 urllib 또는 curl (판별 1회 후 고정)."""
    if not _probe_urllib():
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "replace")
        except Exception as e:
            print(f"  [WARN] urllib 실패, curl 재시도: {str(e)[:60]}")
    try:
        # text=True 는 시스템 기본 인코딩(한국 윈도우=cp949)으로 디코딩해
        # OECD CSV의 UTF-8 문자에서 깨진다 → 바이트로 받아 UTF-8로 직접 디코딩.
        r = subprocess.run(["curl", "-sS", "--max-time", str(timeout), url],
                           capture_output=True, timeout=timeout + 20)
        return r.stdout.decode("utf-8", "replace")
    except Exception as e:
        print(f"  [WARN] 수집 실패: {str(e)[:80]}")
        return ""


def fred(series_id, start="1995-01-01"):
    txt = _http(f"{FRED_CSV}?{urllib.parse.urlencode({'id': series_id, 'cosd': start})}")
    out = {}
    for line in txt.strip().split("\n")[1:]:
        p = line.split(",")
        if len(p) >= 2 and p[1].strip() not in (".", ""):
            try:
                out[p[0].strip()] = float(p[1])
            except ValueError:
                pass
    return out


# ── ECOS (한국은행) ─────────────────────────────────────────────────
def _ecos_key():
    k = os.environ.get("BOK_API_KEY")
    if k:
        return k.strip()
    try:
        from core import get_secret
        return (get_secret("BOK_API_KEY") or "").strip() or None
    except Exception:
        return None


def _ecos(key, path):
    txt = _http(f"{ECOS_BASE}/{path}", timeout=40)
    try:
        return json.loads(txt) if txt else {}
    except Exception:
        return {}


def ecos_find(key, spec):
    """통계표·항목 코드를 이름 키워드로 탐색 → (stat_code, item_code, 라벨) 또는 None.
    코드를 하드코딩하지 않으므로 한국은행이 코드를 바꿔도 계속 동작한다."""
    d = _ecos(key, f"StatisticTableList/{key}/json/kr/1/1000/")
    tables = (d.get("StatisticTableList") or {}).get("row") or []
    if not tables:
        print(f"    [WARN] 통계표 목록 조회 실패 ({spec['label']})")
        return None
    cands = [t for t in tables
             if any(k in (t.get("STAT_NAME") or "") for k in spec["table_kw"]) and t.get("STAT_CODE")]
    print(f"    통계표 후보 {len(cands)}건: " + ", ".join(
        f"{t['STAT_CODE']}({(t.get('STAT_NAME') or '')[:22]})" for t in cands[:4]))
    for t in cands:
        code = t["STAT_CODE"]
        d2 = _ecos(key, f"StatisticItemList/{key}/json/kr/1/500/{code}/")
        items = (d2.get("StatisticItemList") or {}).get("row") or []
        for it in items:
            nm = it.get("ITEM_NAME") or ""
            if spec["item_kw_all"] and not all(k in nm for k in spec["item_kw_all"]):
                continue
            if spec["item_kw_any"] and not any(k in nm for k in spec["item_kw_any"]):
                continue
            print(f"    ✅ 채택 {code} / {it.get('ITEM_CODE')} — {nm[:34]} ({t.get('STAT_NAME','')[:20]})")
            return code, it.get("ITEM_CODE"), nm
    print(f"    [WARN] 조건에 맞는 항목 없음 ({spec['label']})")
    return None


def _ecos_time(t):
    """ECOS 시점 문자열 → ISO 날짜. 주기마다 형식이 달라 방어적으로 판정한다.
        연 '2025' / 분기 '2025Q1' 또는 '20251' / 월 '202501' / 일 '20250131'
    분기 형식은 API 버전에 따라 두 가지가 모두 관측되므로 둘 다 받는다."""
    t = str(t).strip().upper()
    if "Q" in t:
        y, q = t.split("Q")[0], t.split("Q")[1]
        try:
            return f"{y}-{(int(q) - 1) * 3 + 1:02d}-01"
        except ValueError:
            return None
    if not t.isdigit():
        return None
    if len(t) == 4:
        return f"{t}-01-01"
    if len(t) == 5:                       # YYYYQ (분기 축약형)
        return f"{t[:4]}-{(int(t[4]) - 1) * 3 + 1:02d}-01"
    if len(t) == 6:
        return f"{t[:4]}-{t[4:6]}-01"
    if len(t) == 8:
        return f"{t[:4]}-{t[4:6]}-{t[6:8]}"
    return None


def ecos_series(key, stat, item, start="201001", end=None, cycles=("M", "Q", "A")):
    """시계열 → {ISO날짜: value}. 주기를 모르면 월→분기→연 순으로 시도한다.

    ⚠️ 예전엔 월(M)로 고정 호출했는데, 연간 통계(가계금융복지조사 등)에서는
       항목을 제대로 찾고도 빈 응답이 와서 '데이터 없음'으로 오판했다.
       주기가 맞아야 값이 나온다 — 이 폴백이 그 오판을 막는다."""
    now = datetime.now(KST)
    for cyc in cycles:
        if cyc == "A":
            s, e = start[:4], (end or now.strftime("%Y"))[:4]
        elif cyc == "Q":
            s = f"{start[:4]}Q{(int(start[4:6] or 1) - 1) // 3 + 1}" if len(start) >= 6 else f"{start[:4]}Q1"
            e = end or f"{now.year}Q{(now.month - 1) // 3 + 1}"
        else:
            s, e = start, (end or now.strftime("%Y%m"))
        d = _ecos(key, f"StatisticSearch/{key}/json/kr/1/1000/{stat}/{cyc}/{s}/{e}/{item}/")
        rows = (d.get("StatisticSearch") or {}).get("row") or []
        out = {}
        for r in rows:
            t, v = r.get("TIME"), r.get("DATA_VALUE")
            if not t or v in (None, "", "-"):
                continue
            iso = _ecos_time(t)
            if not iso:
                continue
            try:
                out[iso] = float(v)
            except ValueError:
                pass
        if out:
            if cyc != "M":
                print(f"    (주기 {cyc}로 확보 — {len(out)}개)")
            return out
    return {}


# ── 분배·불평등 (ECOS 우선, KOSIS 폴백) ────────────────────────────
KOSIS_BASE = "https://kosis.kr/openapi"
# 가계금융복지조사는 통계청·한국은행·금감원 공동조사라 ECOS에도 실릴 수 있다.
# 새 키(KOSIS) 없이 기존 BOK_API_KEY로 되는지 먼저 시도한다.
ECOS_INEQ = {
    "table_kw": ["가계금융", "가계 금융", "소득분배", "가계자산"],
    "item_kw_all": [],
    "item_kw_any": ["지니", "5분위", "분위", "순자산"],
    "label": "가계 분배지표(지니·분위)",
}


def _kosis_key():
    k = os.environ.get("KOSIS_API_KEY")
    if k:
        return k.strip()
    try:
        from core import get_secret
        return (get_secret("KOSIS_API_KEY") or "").strip() or None
    except Exception:
        return None


def inequality_block():
    """한국 분배지표 — 통화팽창→자산→격차 사슬을 한국 데이터로 검정하기 위한 축.

    미국은 Fed 분배계정(WFRBST01134 등)으로 상위1%/하위50% 자산점유율을 바로 볼 수 있으나
    한국은 동등한 시계열이 FRED에 없다. KOSIS(키 필요) 또는 ECOS(기존 키)로 확보를 시도한다.
    """
    out = {"us": None, "kr": None, "kr_source": None}

    # 1) 미국 — 키 없이 확보되는 기준선. 한국 값이 없어도 비교 축은 남는다.
    top, bot = fred("WFRBST01134", "1999-01-01"), fred("WFRBSB50215", "1999-01-01")
    ks = sorted(set(top) & set(bot))
    if len(ks) >= 8:
        def era(lo, hi):
            s = [k for k in ks if lo <= k[:4] <= hi]
            if len(s) < 4:
                return None
            return {"from": s[0][:7], "to": s[-1][:7],
                    "top1_chg_pp": round(top[s[-1]] - top[s[0]], 1),
                    "bot50_chg_pp": round(bot[s[-1]] - bot[s[0]], 1)}
        out["us"] = {
            "asof": ks[-1][:7], "top1_pct": round(top[ks[-1]], 1), "bot50_pct": round(bot[ks[-1]], 1),
            "ratio": round(top[ks[-1]] / bot[ks[-1]], 1) if bot[ks[-1]] else None,
            "top1_start": round(top[ks[0]], 1), "bot50_start": round(bot[ks[0]], 1), "start": ks[0][:7],
            "eras": {"pre_qe": era("1999", "2008"), "qe": era("2009", "2019"), "post": era("2020", "2026")},
            "source": "Fed Distributional Financial Accounts (FRED WFRBST01134 / WFRBSB50215)",
            "spark_top1": [round(top[k], 1) for k in ks[-60:]],
            "spark_bot50": [round(bot[k], 1) for k in ks[-60:]],
        }
        print(f"  🇺🇸 상위1% {out['us']['top1_pct']}% · 하위50% {out['us']['bot50_pct']}% "
              f"({out['us']['asof']}) · {out['us']['start']} 대비 상위1% "
              f"{out['us']['top1_pct'] - out['us']['top1_start']:+.1f}%p")
    else:
        print("  [WARN] 미국 분배계정 수집 실패")

    # 2) 한국 — ECOS 우선(기존 키), 없으면 KOSIS(신규 키)
    bok = _ecos_key()
    if bok:
        print("  🇰🇷 ECOS에서 분배지표 탐색:")
        f = ecos_find(bok, ECOS_INEQ)
        if f:
            s = ecos_series(bok, f[0], f[1], start="201001")
            if len(s) >= 3:
                kk = sorted(s)
                out["kr"] = {"asof": kk[-1][:7], "value": round(s[kk[-1]], 3),
                             "label": f[2], "series": [{"t": k[:7], "v": round(s[k], 3)} for k in kk[-20:]]}
                out["kr_source"] = f"ECOS {f[0]}/{f[1]}"
                print(f"    ✅ {f[2][:30]} = {s[kk[-1]]} ({kk[-1][:7]})")
    else:
        print("  [INFO] BOK_API_KEY 없음 — ECOS 탐색 생략(로컬)")

    if not out["kr"]:
        kk = _kosis_key()
        if kk:
            print("  🇰🇷 KOSIS 시도…")
            # KOSIS는 통계표ID를 알아야 해서 목록 검색부터 — 실패해도 전체는 계속 진행
            try:
                q = urllib.parse.urlencode({"method": "getList", "apiKey": kk, "vwCd": "MT_ZTITLE",
                                            "parentListId": "A_7", "format": "json", "jsonVD": "Y"})
                d = json.loads(_http(f"{KOSIS_BASE}/statisticsList.do?{q}", timeout=40) or "[]")
                if isinstance(d, dict) and d.get("err"):
                    print(f"    [WARN] KOSIS 오류 {d.get('err')}: {d.get('errMsg')}")
                else:
                    print(f"    KOSIS 목록 {len(d) if isinstance(d, list) else '?'}건 — 통계표 지정 필요")
                    out["kr_source"] = "KOSIS(목록 확보, 통계표 지정 대기)"
            except Exception as e:
                print(f"    [WARN] KOSIS 실패: {str(e)[:60]}")
        else:
            print("  [INFO] KOSIS_API_KEY 없음 — 한국 분배지표 미확보")
    out["note"] = ("미국은 Fed 분배계정으로 자산 점유율을 직접 관측할 수 있으나 한국은 동등한 공개 시계열이 없다. "
                   "따라서 '통화팽창→자산→격차' 사슬을 한국 데이터로 직접 검정하지는 못한 상태이며, "
                   "미국 분포는 참고 기준선일 뿐 한국에 그대로 적용할 수 없다.")
    out["kosis_guide"] = ("KOSIS_API_KEY 발급: kosis.kr → 우측 상단 '오픈API' → 활용신청(무료, 즉시 발급) → "
                          "GitHub 저장소 Settings > Secrets and variables > Actions 에 KOSIS_API_KEY 로 등록")
    return out


def seoul_property_block():
    """서울 아파트 실질 매매가격 = 명목지수 ÷ 소비자물가지수.
    BIS 전국 평균이 감추는 '지역 편차'를 보완한다."""
    key = _ecos_key()
    if not key:
        print("  [INFO] BOK_API_KEY 없음 — 서울 아파트 블록 생략(로컬). GitHub Actions에서는 수집됨.")
        return None
    print("  ECOS 코드 자동 탐색:")
    fa = ecos_find(key, ECOS_TARGETS["seoul_apt"])
    fc = ecos_find(key, ECOS_TARGETS["kr_cpi"])
    if not fa or not fc:
        return None
    apt = ecos_series(key, fa[0], fa[1])
    cpi = ecos_series(key, fc[0], fc[1])
    ks = sorted(set(apt) & set(cpi))
    if len(ks) < 24:
        print(f"  [WARN] 서울 아파트 겹치는 표본 부족 {len(ks)}")
        return None
    base = apt[ks[0]] / cpi[ks[0]]
    real = {k: apt[k] / cpi[k] / base * 100 for k in ks}      # 시작=100 으로 재기준
    last = ks[-1]
    i10 = max(0, len(ks) - 121)                               # 약 10년(120개월) 전
    yv = ks[-13] if len(ks) >= 13 else ks[0]
    out = {
        "name": "서울 아파트", "flag": "🏙️", "asof": last[:7],
        "nominal_index": round(apt[last], 1),
        "real_index": round(real[last], 1), "base": f"{ks[0][:7]}=100 (명목÷CPI)",
        "chg_10y_pct": round((real[last] / real[ks[i10]] - 1) * 100, 1),
        "nominal_10y_pct": round((apt[last] / apt[ks[i10]] - 1) * 100, 1),
        "yoy_pct": round((real[last] / real[yv] - 1) * 100, 1),
        "source": {"apt": f"{fa[0]}/{fa[1]} {fa[2][:30]}", "cpi": f"{fc[0]}/{fc[1]} {fc[2][:20]}"},
        "spark": [round(real[k], 1) for k in ks[-60:]],
        "note": "ECOS 자동 탐색으로 통계표·항목 코드를 찾아 수집(코드 하드코딩 없음).",
    }
    print(f"  🏙️ 서울 아파트: 실질 {out['real_index']} · 10년 실질 {out['chg_10y_pct']:+.1f}% "
          f"(명목 {out['nominal_10y_pct']:+.1f}%)")
    return out


# ── 품목별 물가·임금 (체감물가 계산기/임금협상 근거용) ──────────────
OECD_PRICES = ("https://sdmx.oecd.org/public/rest/data/OECD.SDD.TPS,DSD_PRICES@DF_PRICES_ALL,"
               "/all?startPeriod={start}&format=csvfilewithlabels")
# COICOP 영문 항목명 → 화면 표기. 계산기에서 사용자가 비중을 조정할 축이다.
CPI_ITEMS = [
    ("Food and non-alcoholic beverages", "식료품·비주류음료", 13.0),
    ("Housing, water, electricity, gas and other fuels", "주거·수도·광열", 17.0),
    ("Transport", "교통", 11.0),
    ("Fuels and lubricants for personal transport equipment", "자동차 연료", 0.0),
    ("Restaurants and hotels", "음식·숙박", 13.0),
    ("Recreation and culture", "오락·문화", 7.0),
    ("Communication", "통신", 5.0),
    ("Health", "보건", 8.0),
    ("Education", "교육", 7.0),
    ("Clothing and footwear", "의류·신발", 6.0),
    ("Furnishings, household equipment and routine household maintenance", "가정용품·가사서비스", 5.0),
    ("Alcoholic beverages, tobacco and narcotics", "주류·담배", 2.0),
    ("Miscellaneous goods and services", "기타 상품·서비스", 6.0),
]
WAGE_KR = "LCEAMN01KRQ661S"      # 한국 임금지수(분기). 상여 변동이 커 4분기 이동평균으로 평활


def cpi_items_block():
    """한국 품목별 전년동월비 — 체감물가 계산기의 입력값."""
    txt = _http(OECD_PRICES.format(start="2025-06"), timeout=180)
    if not txt or "REF_AREA" not in txt:
        print("  [WARN] OECD 물가 수집 실패")
        return None
    latest, vals = "", {}
    for r in csv.DictReader(io.StringIO(txt)):
        if (r.get("REF_AREA") != "KOR" or r.get("FREQ") != "M"
                or r.get("Measure") != "Consumer price index"
                or r.get("Transformation") != "Growth rate, over 1 year"):
            continue
        t = r.get("TIME_PERIOD") or ""
        try:
            vals.setdefault(t, {})[r["Expenditure"]] = float(r["OBS_VALUE"])
        except (ValueError, KeyError):
            pass
        latest = max(latest, t)
    if not latest:
        return None
    cur = vals[latest]
    items = [{"key": en, "name": ko, "default_weight": w, "yoy_pct": round(cur[en], 2)}
             for en, ko, w in CPI_ITEMS if en in cur]
    total = cur.get("Total")
    print(f"  품목 {len(items)}개 ({latest}) · 총지수 {total}%")
    for it in sorted(items, key=lambda x: -x["yoy_pct"])[:4]:
        print(f"    {it['name']} {it['yoy_pct']:+.2f}%")
    return {"asof": latest, "total_yoy_pct": round(total, 2) if total is not None else None,
            "items": items,
            "note": ("가중치는 한국 CPI 실제 가중치의 근사값(기본값)이며 사용자가 조정하는 축이다. "
                     "'자동차 연료'는 교통에 포함된 세부항목이라 기본 가중 0 — 차를 많이 쓰는 경우만 올려 쓴다.")}


def wage_block():
    """한국 명목임금 상승률(4분기 이동평균) — 실질임금 계산의 분자."""
    d = fred(WAGE_KR, "2013-01-01")
    ks = sorted(d)
    if len(ks) < 9:
        print("  [WARN] 임금지수 표본 부족")
        return None
    ma = [(ks[i], sum(d[k] for k in ks[i - 3:i + 1]) / 4) for i in range(3, len(ks))]
    series = []
    for i in range(4, len(ma)):
        series.append({"q": ma[i][0][:7], "yoy_pct": round((ma[i][1] / ma[i - 4][1] - 1) * 100, 2)})
    last = series[-1]
    print(f"  명목임금 {last['q']} {last['yoy_pct']:+.2f}% (4분기 이동평균)")
    return {"asof": last["q"], "nominal_yoy_pct": last["yoy_pct"], "series": series[-24:],
            "series_id": WAGE_KR,
            "method": "분기 임금지수의 4분기 이동평균 전년비 — 상여금·계절 변동을 평활",
            "note": "실질임금 = 명목임금 상승률 − 물가상승률. 어떤 물가를 쓰느냐에 따라 결과가 달라진다."}


# ── 노동 몫 (피용자보수 / GDP) ──────────────────────────────────────
# 임금단협의 가장 오래된 거시 논거: "생산된 소득 중 노동에 돌아가는 몫이 얼마인가".
# 한국은행도 같은 개념을 공표하지만(2022년부터 '노동소득분배율' → '피용자보수비율'로 개칭)
# ECOS는 API 키가 필요하다. OECD 국민계정은 키 없이 4개국 25년치를 주므로 이쪽을 쓴다.
#
# ⚠️ 분모가 다르면 값이 크게 달라진다 — 아래 값은 한국은행 발표치와 다르다.
#    · 여기(OECD)   = 피용자보수 ÷ GDP                    → 한국 48% 수준
#    · 한국은행     = 피용자보수 ÷ 요소비용국민소득       → 한국 68% 수준
#    요소비용국민소득은 GDP에서 고정자본소모와 생산세를 뺀 것이라 분모가 작다.
#    둘 다 맞는 계산이며, 국제 비교에는 분모가 통일된 전자를 쓴다.
OECD_NAMAIN = ("https://sdmx.oecd.org/public/rest/data/OECD.SDD.NAD,DSD_NAMAIN10@DF_TABLE1,"
               "/A.{areas}.S1..D1+B1GQ.......?startPeriod=2000&format=csvfilewithlabels")
LABOR_AREAS = [("KOR", "한국", "🇰🇷"), ("USA", "미국", "🇺🇸"),
               ("JPN", "일본", "🇯🇵"), ("DEU", "독일", "🇩🇪")]


def labor_share_block():
    """피용자보수/GDP — 국민경제 전체에서 노동에 돌아가는 몫."""
    url = OECD_NAMAIN.format(areas="+".join(a for a, _, _ in LABOR_AREAS))
    txt = _http(url, timeout=120)
    if not txt:
        print("  [WARN] OECD 국민계정 조회 실패")
        return None
    # 같은 거래코드가 산업별(ACTIVITY)·표별(TABLE_IDENTIFIER)로 중복 수록되므로
    # 전산업(_T)·경상가격(V)·자국통화(XDC)·T0103 한 조합만 남긴다. 안 걸러내면 값이 뒤섞인다.
    got = {}
    for r in csv.DictReader(io.StringIO(txt)):
        if (r.get("UNIT_MEASURE") != "XDC" or r.get("PRICE_BASE") != "V"
                or r.get("TRANSFORMATION") != "N" or r.get("SECTOR") != "S1"
                or r.get("ACTIVITY") not in ("_T", "_Z") or r.get("TABLE_IDENTIFIER") != "T0103"):
            continue
        try:
            got.setdefault((r["REF_AREA"], r["TRANSACTION"]), {})[r["TIME_PERIOD"]] = float(r["OBS_VALUE"])
        except (ValueError, KeyError):
            pass
    out = []
    for code, name, flag in LABOR_AREAS:
        d1, gq = got.get((code, "D1"), {}), got.get((code, "B1GQ"), {})
        ys = sorted(set(d1) & set(gq))
        if len(ys) < 10:
            print(f"  [WARN] {name} 표본 부족 {len(ys)}")
            continue
        sh = {y: d1[y] / gq[y] * 100 for y in ys if gq[y]}
        last, first = ys[-1], ys[0]
        y10 = ys[-11] if len(ys) >= 11 else first
        out.append({
            "code": code, "name": name, "flag": flag, "asof": last,
            "share_pct": round(sh[last], 1),
            "start_year": first, "start_pct": round(sh[first], 1),
            "chg_since_start_pp": round(sh[last] - sh[first], 1),
            "chg_10y_pp": round(sh[last] - sh[y10], 1),
            "series": [{"y": y, "v": round(sh[y], 1)} for y in ys],
        })
        print(f"  {flag} {name} {sh[last]:.1f}% ({last}) · {first} 대비 {sh[last] - sh[first]:+.1f}%p "
              f"· 10년 {sh[last] - sh[y10]:+.1f}%p")
    if not out:
        return None
    return {
        "areas": out,
        "definition": "피용자보수(D1) ÷ 국내총생산(B1GQ) — 경상가격·자국통화·전산업",
        "source": "OECD SDMX DSD_NAMAIN10@DF_TABLE1 (키 불필요)",
        "bok_note": ("한국은행은 같은 개념을 '피용자보수비율'(2022년까지 '노동소득분배율')로 공표하며 "
                     "분모가 요소비용국민소득이라 값이 더 크다(한국 68% 수준). 분모만 다를 뿐 둘 다 맞는 계산이고, "
                     "국제 비교에는 분모가 통일된 GDP 기준을 쓴다."),
        "caveat": ("자영업자의 노동소득은 피용자보수가 아니라 영업잉여·혼합소득에 들어간다. "
                   "자영업 비중이 큰 한국에서는 이 지표가 실제 노동 몫을 과소평가한다 — "
                   "한국은행이 '노동소득분배율'이라는 이름을 버린 이유이기도 하다. "
                   "자영업 비중이 줄어들면 지표는 저절로 올라가므로 상승분 전부를 분배 개선으로 읽으면 안 된다."),
    }


# ── 가계 소비 여력 (1인당 실질 가계소비) ───────────────────────────
# 실질임금이 올랐다는 통계와 "쓸 돈이 없다"는 체감이 어긋날 때, 그 사이를 메우는 지표.
# 임금은 근로자 1인 기준이지만 소비는 인구 1인 기준이라 가구 구성·고용 변화까지 반영된다.
#
# ⚠️ 한국은 OECD 가계 대시보드에 '가처분소득'과 '저축률'이 수록돼 있지 않다(2026-08 확인).
#    수록된 나라는 있으나 한국은 소비·실업률·소비자심리만 제공된다.
#    가계총처분가능소득·순저축률이 필요하면 한국은행 ECOS(키 필요) 경로를 따로 써야 한다.
OECD_HHDASH = ("https://sdmx.oecd.org/public/rest/data/OECD.SDD.NAD,DSD_HHDASH@DF_HHDASH_CTRY,"
               "/Q.{areas}.P3S1M_R_POP_GR.PC?startPeriod=2015&format=csvfilewithlabels")
HH_AREAS = [("KOR", "한국", "🇰🇷"), ("USA", "미국", "🇺🇸"), ("JPN", "일본", "🇯🇵")]


def household_block():
    """1인당 실질 가계소비 증가율(전년동기비, 분기)."""
    txt = _http(OECD_HHDASH.format(areas="+".join(a for a, _, _ in HH_AREAS)), timeout=90)
    if not txt:
        print("  [WARN] OECD 가계 대시보드 조회 실패")
        return None
    got = {}
    for r in csv.DictReader(io.StringIO(txt)):
        try:
            got.setdefault(r["REF_AREA"], {})[r["TIME_PERIOD"]] = float(r["OBS_VALUE"])
        except (ValueError, KeyError):
            pass
    out = []
    for code, name, flag in HH_AREAS:
        d = got.get(code) or {}
        qs = sorted(q for q in d if "-Q" in q)
        if len(qs) < 4:
            print(f"  [WARN] {name} 분기 표본 부족 {len(qs)}")
            continue
        last = qs[-1]
        y4 = sum(d[q] for q in qs[-4:]) / 4          # 최근 4분기 평균 = 연간 체감에 가깝다
        out.append({"code": code, "name": name, "flag": flag, "asof": last,
                    "yoy_pct": round(d[last], 2), "avg4q_pct": round(y4, 2),
                    "series": [{"q": q, "v": round(d[q], 2)} for q in qs[-16:]]})
        print(f"  {flag} {name} {last} {d[last]:+.2f}% · 최근 4분기 평균 {y4:+.2f}%")
    if not out:
        return None
    return {
        "areas": out,
        "definition": "1인당 실질 가계·비영리단체 최종소비지출 증가율(전년동기비)",
        "source": "OECD SDMX DSD_HHDASH@DF_HHDASH_CTRY (키 불필요)",
        "why": ("실질임금은 '근로자 1인'이지만 이 지표는 '인구 1인'이라 가구원 수·고용률 변화까지 반영한다. "
                "실질임금이 플러스인데 이 값이 0 부근이면, 오른 임금이 늘어난 부양 부담이나 "
                "고용 구성 변화로 흡수됐다는 뜻이다."),
        "caveat": ("소비는 저축을 헐거나 빚을 내서도 늘릴 수 있으므로 소득의 대용치로 쓰면 안 된다. "
                   "한국은 OECD 대시보드에 가처분소득·저축률이 수록돼 있지 않아 소비만 본다."),
    }


# ── 통계 도우미 ─────────────────────────────────────────────────────
def yoy(d, per):
    ks = sorted(d)
    return {k: (d[k] / d[ks[i - per]] - 1) * 100
            for i, k in enumerate(ks) if i >= per and d[ks[i - per]]}


def to_q(d):
    """월별 키 → 분기 첫 달 키."""
    out = {}
    for k, v in d.items():
        m = int(k[5:7])
        out[f"{k[:4]}-{((m - 1) // 3) * 3 + 1:02d}-01"] = v
    return out


def shift_m(d, months):
    """관측을 months만큼 미래로 이동 = 'months 선행' 검정."""
    out = {}
    for k, v in d.items():
        y, m = int(k[:4]), int(k[5:7])
        m += months
        y += (m - 1) // 12
        m = (m - 1) % 12 + 1
        out[f"{y:04d}-{m:02d}-01"] = v
    return out


def corr(a, b, minn=12):
    ks = sorted(set(a) & set(b))
    if len(ks) < minn:
        return None, len(ks)
    x = [a[k] for k in ks]
    y = [b[k] for k in ks]
    n = len(x)
    mx, my = sum(x) / n, sum(y) / n
    sx = sum((v - mx) ** 2 for v in x) ** 0.5
    sy = sum((v - my) ** 2 for v in y) ** 0.5
    if not sx or not sy:
        return None, n
    return sum((x[i] - mx) * (y[i] - my) for i in range(n)) / (sx * sy), n


def best_lag(src, dst, lags, minn=12):
    best = None
    grid = []
    for lg in lags:
        r, n = corr(shift_m(src, lg), dst, minn)
        if r is None:
            continue
        grid.append({"lag_m": lg, "r": round(r, 3), "n": n})
        if best is None or abs(r) > abs(best["r"]):
            best = {"lag_m": lg, "r": round(r, 3), "n": n}
    return best, grid


# ── 블록 1: 주요국 통화량 ───────────────────────────────────────────
def money_block():
    txt = _http(OECD_MONAGG.format(start="2015-01"), timeout=180)
    if not txt or "REF_AREA" not in txt:
        print("  [WARN] OECD 통화총량 수집 실패")
        return []
    series = {}
    for r in csv.DictReader(io.StringIO(txt)):
        if (r.get("Measure") != "M3" or r.get("FREQ") != "M"
                or r.get("UNIT_MEASURE") != "XDC"):
            continue
        area = r.get("REF_AREA")
        try:
            series.setdefault(area, {})[r["TIME_PERIOD"] + "-01"] = float(r["OBS_VALUE"])
        except (ValueError, KeyError):
            pass

    fx = {}
    try:
        import yfinance as yf
        for cur, tk in FX_TICKER.items():
            try:
                h = yf.Ticker(tk).history(period="5d", interval="1d")["Close"].dropna()
                if len(h):
                    fx[cur] = float(h.iloc[-1])
            except Exception:
                pass
    except Exception:
        print("  [WARN] yfinance 없음 — 달러 환산 비중 생략")

    out = []
    for code, name, flag, cur in AREAS:
        d = series.get(code) or {}
        if len(d) < 14:
            print(f"  [WARN] {name} 표본 부족 {len(d)}")
            continue
        g = yoy(d, 12)
        ks = sorted(d)
        asof, last = ks[-1], d[ks[-1]]
        # 달러 환산(백만 자국통화 → 십억 달러). EURUSD만 곱셈, 나머지는 나눗셈.
        usd = None
        if cur == "USD":
            usd = last / 1000
        elif cur == "EUR" and fx.get("EUR"):
            usd = last * fx["EUR"] / 1000
        elif fx.get(cur):
            usd = last / fx[cur] / 1000
        gk = sorted(g)
        # 연도별 증가율(각 연도 마지막 관측의 전년동월비) — 단협·장기추세용
        annual = []
        for y in range(int(gk[0][:4]) if gk else 2016, int(gk[-1][:4]) + 1 if gk else 2026):
            yy = [k for k in gk if k.startswith(str(y))]
            if yy:
                annual.append({"year": y, "yoy_pct": round(g[yy[-1]], 2), "month": yy[-1][:7]})
        out.append({
            "annual": annual,
            "code": code, "name": name, "flag": flag, "currency": cur,
            "asof": asof[:7], "level_local_mn": round(last, 0),
            "level_usd_bn": round(usd, 1) if usd else None,
            "yoy_pct": round(g[gk[-1]], 2) if gk else None,
            "yoy_1y_ago": round(g[gk[-13]], 2) if len(gk) >= 13 else None,
            "spark": [round(g[k], 2) for k in gk[-36:]],
        })
        print(f"  {flag} {name}: M3 {last:,.0f}백만{cur} ({asof[:7]}) "
              f"YoY {g[gk[-1]]:+.2f}%" + (f" · ${usd:,.0f}B" if usd else ""))
    tot = sum(x["level_usd_bn"] for x in out if x["level_usd_bn"])
    for x in out:
        x["share_pct"] = round(x["level_usd_bn"] / tot * 100, 1) if (tot and x["level_usd_bn"]) else None
    return out


# ── 블록 2: 전달경로 ────────────────────────────────────────────────
def transmission_block():
    m2 = yoy(fred("M2SL"), 12)
    cpi = yoy(fred("CPIAUCSL"), 12)
    tax = yoy(fred("W006RC1Q027SBEA"), 4)          # 연방 총수입(분기)
    lags = [0, 3, 6, 9, 12, 15, 18, 21, 24]
    b_cpi, g_cpi = best_lag(m2, cpi, lags)
    b_tax, g_tax = best_lag(to_q(m2), tax, lags, minn=8)
    b_ct, g_ct = best_lag(to_q(cpi), tax, lags, minn=8)

    eras = []
    if b_cpi:
        for lo, hi, lab in [("1996", "2007", "1996~2007"), ("2008", "2019", "2008~2019 (QE기)"),
                            ("2020", "2026", "2020~2026 (팬데믹 후)")]:
            a = {k: v for k, v in shift_m(m2, b_cpi["lag_m"]).items() if lo <= k[:4] <= hi}
            b = {k: v for k, v in cpi.items() if lo <= k[:4] <= hi}
            r, n = corr(a, b)
            if r is not None:
                eras.append({"label": lab, "r": round(r, 3), "n": n})
    return {
        "m2_to_cpi": {"best": b_cpi, "grid": g_cpi, "eras": eras},
        "m2_to_tax": {"best": b_tax, "grid": g_tax},
        "cpi_to_tax": {"best": b_ct, "grid": g_ct},
        "verdict": ("통화량은 세수에 직접 닿지 않는다 — 물가를 매개로 전달된다. "
                    "M2→물가 상관이 M2→세수보다 크고, 물가→세수는 시차 없이 동행한다."),
        "caveat": ("구간별로 부호가 뒤집힌다(QE기엔 역상관). '돈을 풀면 물가가 오른다'는 "
                   "상시 법칙이 아니라 국면 의존적이다. 상관은 인과가 아니다."),
    }


# ── 블록 3: 실질주택가격 ────────────────────────────────────────────
def property_block():
    out = []
    m2q = to_q(yoy(fred("M2SL"), 12))
    for sid, name, flag in PROPERTY:
        d = fred(sid)
        if len(d) < 20:
            continue
        ks = sorted(d)
        g = yoy(d, 4)
        b, _ = best_lag(m2q, g, [0, 3, 6, 9, 12, 15, 18, 21, 24], minn=8)
        i10 = max(0, len(ks) - 41)
        out.append({
            "name": name, "flag": flag, "series_id": sid, "asof": ks[-1][:7],
            "index": round(d[ks[-1]], 1), "base": "2010=100",
            "chg_10y_pct": round((d[ks[-1]] / d[ks[i10]] - 1) * 100, 1),
            "yoy_pct": round(g[sorted(g)[-1]], 1) if g else None,
            "m2_lead": b,
            "spark": [round(d[k], 1) for k in ks[-40:]],
        })
        print(f"  {flag} {name}: {d[ks[-1]]:.1f} ({ks[-1][:7]}) · 10년 "
              f"{(d[ks[-1]] / d[ks[i10]] - 1) * 100:+.1f}%")
    return out


# ── 블록 4: 재정(관세·세수) ─────────────────────────────────────────
def fiscal_block():
    duty = fred("B235RC1Q027SBEA")      # 관세 등 수입 관련 세금(분기·연율 십억$)
    tax = fred("W006RC1Q027SBEA")       # 연방 총수입(분기·연율 십억$)
    dxy = fred("DTWEXBGS")              # 광의 달러지수(일별)
    dq = {}
    for k, v in dxy.items():
        m = int(k[5:7])
        dq.setdefault(f"{k[:4]}-{((m - 1) // 3) * 3 + 1:02d}-01", []).append(v)
    dqa = {k: sum(v) / len(v) for k, v in dq.items()}
    rows = []
    for k in sorted(set(duty) & set(tax))[-10:]:
        rows.append({"q": k[:7], "duty_bn": round(duty[k], 1), "tax_bn": round(tax[k], 1),
                     "duty_share_pct": round(duty[k] / tax[k] * 100, 2) if tax[k] else None,
                     "dxy": round(dqa[k], 1) if k in dqa else None})
    if not rows:
        return None
    first, last = rows[0], rows[-1]
    print(f"  관세수입 {first['q']} {first['duty_bn']} → {last['q']} {last['duty_bn']}십억$ "
          f"(총세수 대비 {first['duty_share_pct']}% → {last['duty_share_pct']}%)")
    return {"rows": rows,
            "note": ("관세수입=연방 '생산·수입세' 중 관세(BEA, 분기 연율). 총세수 대비 비중과 "
                     "달러지수를 나란히 둬 '강달러가 관세 부담을 흡수했는가'를 눈으로 확인할 수 있게 함.")}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    now = datetime.now(KST)
    print("=" * 62)
    print("  통화량·물가·자산 전달경로 모니터")
    print("=" * 62)

    print("\n[1] 주요 5개 경제권 통화량 (OECD M3)")
    money = money_block()
    print("\n[2] 전달경로 (미국 M2 → 물가 → 세수)")
    trans = transmission_block()
    if trans["m2_to_cpi"]["best"]:
        b = trans["m2_to_cpi"]["best"]
        print(f"  M2→CPI 최강 {b['lag_m']}개월 r={b['r']} (n={b['n']})")
        for e in trans["m2_to_cpi"]["eras"]:
            print(f"    {e['label']:<22} r={e['r']:+.3f} (n={e['n']})")
    for k, lab in (("m2_to_tax", "M2→세수"), ("cpi_to_tax", "CPI→세수")):
        b = trans[k]["best"]
        if b:
            print(f"  {lab} 최강 {b['lag_m']}개월 r={b['r']} (n={b['n']})")
    print("\n[3] 실질주택가격 (BIS 4개국 + 서울 아파트)")
    prop = property_block()
    seoul = seoul_property_block()
    print("\n[4] 재정 (관세·세수)")
    fisc = fiscal_block()
    print("\n[5] 품목별 물가 (체감물가 계산기 입력)")
    cpi_items = cpi_items_block()
    print("\n[6] 임금 (실질임금·단협 근거)")
    wage = wage_block()
    print("\n[7] 분배·불평등 (통화팽창→자산→격차 검정축)")
    ineq = inequality_block()
    print("\n[8] 노동 몫 (피용자보수/GDP — 임금단협 거시 기준선)")
    lshare = labor_share_block()
    print("\n[9] 가계 소비 여력 (1인당 실질 가계소비)")
    hh = household_block()

    if not money and not prop:
        print("\n[ERROR] 주요 블록 수집 실패 — 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "money": money, "transmission": trans, "property": prop,
        "seoul_property": seoul, "fiscal": fisc,
        "cpi_items": cpi_items, "wage": wage, "inequality": ineq, "labor_share": lshare, "household": hh,
        "sources": {
            "money": "OECD SDMX DF_MONAGG (M3, 월별, 자국통화) — 키 불필요",
            "us_macro": "FRED M2SL·CPIAUCSL·W006RC1Q027SBEA·B235RC1Q027SBEA·DTWEXBGS",
            "property": "BIS 실질주거용부동산가격(FRED 경유, 2010=100, 명목÷물가)",
            "fx": "yfinance 최신 환율(비중 계산용 근사)",
            "seoul": "한국은행 ECOS — 통계표·항목 코드를 이름 키워드로 자동 탐색(하드코딩 없음)",
            "labor_share": "OECD SDMX DSD_NAMAIN10@DF_TABLE1 — 피용자보수÷GDP, 키 불필요",
            "household": "OECD SDMX DSD_HHDASH@DF_HHDASH_CTRY — 1인당 실질 가계소비 증가율, 키 불필요",
        },
        "caveats": [
            "통화량 '증가율'은 자국통화 기준이라 환율 영향이 없지만, '비중'은 최신 환율로 환산한 근사치다(과거 환율 미반영).",
            "5개 경제권은 세계 통화량의 큰 부분이지만 전부가 아니다 — '세계 비중'이 아니라 '5개 경제권 내 비중'이다.",
            ("중국 M3가 미국보다 큰 것은 '돈을 더 많이 풀어서'가 아니라 금융구조 차이다. "
             "중국은 가계·기업 자금이 은행예금에 몰려 있어 광의통화/GDP가 구조적으로 200%대인 반면, "
             "미국은 MMF·채권 등 은행 밖 자산 비중이 커서 90%대다. 절대 규모 비교보다 '증가율'과 "
             "'자국 내 추세'를 보는 편이 타당하다."),
            "실질주택가격은 전국 평균이라 서울·수도권 등 특정 지역 체감과 크게 다를 수 있다.",
            "상관관계는 인과가 아니며, 구간을 나누면 부호가 뒤집히는 관계가 있다.",
        ],
        "note": "공개 데이터(OECD·FRED·BIS)의 규칙기반 요약 · 투자자문 아님",
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"\n[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
