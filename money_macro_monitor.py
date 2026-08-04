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


def ecos_series(key, stat, item, start="201001", end=None):
    """월별 시계열 → {YYYY-MM-01: value}"""
    end = end or datetime.now(KST).strftime("%Y%m")
    d = _ecos(key, f"StatisticSearch/{key}/json/kr/1/1000/{stat}/M/{start}/{end}/{item}/")
    rows = (d.get("StatisticSearch") or {}).get("row") or []
    out = {}
    for r in rows:
        t, v = r.get("TIME"), r.get("DATA_VALUE")
        if not t or v in (None, "", "-"):
            continue
        try:
            out[f"{t[:4]}-{t[4:6]}-01"] = float(v)
        except ValueError:
            pass
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
        out.append({
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

    if not money and not prop:
        print("\n[ERROR] 주요 블록 수집 실패 — 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "money": money, "transmission": trans, "property": prop,
        "seoul_property": seoul, "fiscal": fisc,
        "sources": {
            "money": "OECD SDMX DF_MONAGG (M3, 월별, 자국통화) — 키 불필요",
            "us_macro": "FRED M2SL·CPIAUCSL·W006RC1Q027SBEA·B235RC1Q027SBEA·DTWEXBGS",
            "property": "BIS 실질주거용부동산가격(FRED 경유, 2010=100, 명목÷물가)",
            "fx": "yfinance 최신 환율(비중 계산용 근사)",
            "seoul": "한국은행 ECOS — 통계표·항목 코드를 이름 키워드로 자동 탐색(하드코딩 없음)",
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
