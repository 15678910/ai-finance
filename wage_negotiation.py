"""회사별 임금협상 근거자료 — DART 직원현황 + 손익 (지불능력 분석)

왜 만들었나
----------
임금 단체협약에서 노동측은 '생계비(물가)'를, 사측은 '지불능력'을 근거로 든다.
체감물가 계산기가 앞의 절반을 채웠으므로, 뒤의 절반인 지불능력을 붙인다.
한쪽 편을 들지 않고 **양쪽 근거를 같은 화면에 놓는 것**이 목적이다.

DART에서 자동으로 가져오는 것 (상장사 사업보고서 공시 의무 항목)
  · empSttus     — 직원 수, 연간급여 총액, 1인평균 급여액
  · fnlttSinglAcntAll — 매출액, 영업이익

여기서 계산하는 것
  · 인건비/매출 비중 — 인상 여력의 핵심 지표
  · 영업이익률, 1인당 매출
  · 인상률 시나리오 — X% 인상 시 인건비 증가액과 영업이익률 변화

⚠️ 한계
  · 상장사만 가능(비상장은 공시 의무 없음) → 화면에서 수동 입력도 지원
  · 공시 급여총액은 등기임원 보수 등 포함 범위가 회사마다 다를 수 있다
  · 연결/별도 기준 차이로 매출과 급여의 기준연도·범위가 어긋날 수 있다
  → 협상 자료로 쓸 때는 반드시 원문 사업보고서로 교차 확인할 것

출력: docs/wage_negotiation.json
🚨 어느 쪽 주장도 대변하지 않는 계산 결과 · 노무·법률 자문 아님.
"""

import json
import os
import sys
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "wage_negotiation.json")
DART = "https://opendart.fss.or.kr/api"
UA = {"User-Agent": "Mozilla/5.0 (compatible; ai-finance-wage)"}

# 관심종목 — 사용자가 소속·비교하려는 회사가 여기 없으면 화면에서 직접 입력한다.
WATCH = [
    ("000660", "SK하이닉스"), ("005930", "삼성전자"), ("108490", "로보티즈"),
    ("003550", "LG"), ("066570", "LG전자"), ("042700", "한미반도체"),
    ("009150", "삼성전기"), ("373220", "LG에너지솔루션"),
]
YEARS = 3          # 최근 3개 사업연도까지 시도


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


def _num(s):
    try:
        v = str(s).replace(",", "").strip()
        return None if v in ("", "-", "None") else float(v)
    except Exception:
        return None


def corp_map(key):
    """종목코드 → DART 고유번호. dart_financials.py와 동일 방식."""
    import io
    import re
    import zipfile
    raw = _get(f"{DART}/corpCode.xml?crtfc_key={key}", timeout=60)
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        xml = z.read(z.namelist()[0]).decode("utf-8", "replace")
    m = {}
    for blk in re.findall(r"<list>(.*?)</list>", xml, re.S):
        sc = re.search(r"<stock_code>\s*(\S+)\s*</stock_code>", blk)
        cc = re.search(r"<corp_code>\s*(\S+)\s*</corp_code>", blk)
        if sc and cc and sc.group(1).strip():
            m[sc.group(1).strip()] = cc.group(1).strip()
    return m


TOTAL_KW = ("합계", "소계", "총계", "전체")


def _detail_rows(rows):
    """부문별 상세 행만 남긴다.

    ⚠️ 회사에 따라 '성별합계' 같은 소계 행을 상세 행과 **함께** 공시한다.
       그대로 다 더하면 인원·급여가 정확히 2배가 된다.
       (실측 2026-08-05: 삼성전자 128,881명 → 257,762명, LG전자·삼성전기도 동일 증상)
       소계 행만 공시한 회사도 있으므로, 상세 행이 없을 때만 소계를 쓴다.
    """
    detail = [r for r in rows
              if not any(k in (r.get("fo_bbm") or "") for k in TOTAL_KW)]
    return detail or rows


def _emp_split(r):
    """행 하나의 (정규직, 기간제) 인원 → 판별 불가면 None.

    ⚠️ 단시간 근로자(*_abacpt_labrr_co)를 본 인원에 포함해 공시하는 회사와
       따로 떼어 공시하는 회사가 섞여 있다. 문서만 보고는 알 수 없으므로
       그 행의 합계(sm)와 맞춰 보아 어느 방식인지 판별한다.
         삼성전자      rgllbr + cnttk           = sm  → 단시간이 이미 포함
         LG에너지솔루션 rgllbr+단시간 + cnttk+단시간 = sm  → 단시간이 별도
    """
    sm = _num(r.get("sm")) or 0
    rg = _num(r.get("rgllbr_co")) or 0
    ra = _num(r.get("rgllbr_abacpt_labrr_co")) or 0
    ct = _num(r.get("cnttk_co")) or 0
    ca = _num(r.get("cnttk_abacpt_labrr_co")) or 0
    if not sm or not (rg or ct):
        return None
    if abs(rg + ct - sm) < 1:
        return rg, ct
    if abs(rg + ra + ct + ca - sm) < 1:
        return rg + ra, ct + ca
    return None


def employees(key, corp, year):
    """직원 현황 — 인원·급여 + 고용형태(정규직/기간제) 구분 (사업보고서 11011).

    DART는 정규직(rgllbr_co)과 기간제(cnttk_co) '인원'은 구분 공시하지만
    급여는 합산으로만 공시한다. 따라서 고용형태별 임금 격차는 이 API로 알 수 없고,
    화면에서 사용자가 직접 입력해야 한다(단체협약 자료·급여명세로 확인 가능).

    ⚠️ rgllbr_abacpt_labrr_co(정규직 단시간)는 rgllbr_co에 **이미 포함된 부분집합**이다.
       더하면 이중 계상된다. 실측으로 확인: 삼성전자 rgllbr_co + cnttk_co = sm 과 정확히 일치.
    """
    try:
        d = json.loads(_get(f"{DART}/empSttus.json?crtfc_key={key}&corp_code={corp}"
                            f"&bsns_year={year}&reprt_code=11011").decode("utf-8", "replace"))
    except Exception:
        return None
    if d.get("status") != "000" or not d.get("list"):
        return None
    rows_all = d["list"]
    # 인원·고용형태·근속은 상세 행에서
    n, reg, tmp, tenure_w, tenure_n, split_ok = 0, 0, 0, 0.0, 0, True
    for r in _detail_rows(rows_all):
        c = _num(r.get("sm"))                       # 직원 수 합계
        tn = _num(r.get("avrg_cnwk_sdytrn"))        # 평균 근속연수
        if c:
            n += int(c)
        sp = _emp_split(r)
        if sp is None:
            split_ok = False
        else:
            reg += int(sp[0])
            tmp += int(sp[1])
        if tn and c:
            tenure_w += tn * c
            tenure_n += int(c)
    if not n:
        return None

    # 급여는 소계 행에만 싣는 회사가 있다(삼성전자 등) → 급여가 있는 행만 골라 다시 중복 제거.
    # 상세·소계 어느 쪽에 있든 한 벌만 더해진다.
    pay, avg_w, avg_n = 0.0, 0.0, 0
    for r in _detail_rows([r for r in rows_all
                           if _num(r.get("fyer_salary_totamt")) or _num(r.get("jan_salary_am"))]):
        c = _num(r.get("sm"))
        t = _num(r.get("fyer_salary_totamt"))       # 연간급여 총액
        a = _num(r.get("jan_salary_am"))            # 1인평균 급여액
        if t:
            pay += t
        if a and c:                                 # 부문별 평균을 인원 가중으로 합산
            avg_w += a * c
            avg_n += int(c)

    typed = reg + tmp
    # 검산 — 고용형태 합이 인원 합계와 어긋나면 공시 구조가 예상과 다른 것이다.
    # 조용히 틀린 숫자를 내보내느니 경고를 남기고 고용형태만 버린다.
    if not split_ok or (typed and abs(typed - n) > max(2, n * 0.01)):
        if typed:
            print(f"      [WARN] 고용형태 합 {typed:,} ≠ 인원 {n:,} — 고용형태 제외")
        reg = tmp = typed = 0
    return {"headcount": n,
            "payroll_total": round(pay) if pay else None,
            "avg_pay": round(avg_w / avg_n) if avg_n else (round(pay / n) if pay else None),
            "regular": reg or None, "temporary": tmp or None,
            "temp_ratio_pct": round(tmp / typed * 100, 1) if typed else None,
            "avg_tenure_yr": round(tenure_w / tenure_n, 1) if tenure_n else None}


def financials(key, corp, year):
    """매출액·영업이익·당기순이익 + 전년 동일 항목 (연결 우선, 없으면 별도).

    DART 응답은 당기(thstrm)와 전기(frmtrm)를 함께 주므로 한 번의 호출로 증가율까지 낼 수 있다.
    성과배분형 임금 산식은 '전년 대비 이익 증가율'을 쓰므로 전기 값이 필수다.
    """
    for fs in ("CFS", "OFS"):
        try:
            d = json.loads(_get(f"{DART}/fnlttSinglAcntAll.json?crtfc_key={key}&corp_code={corp}"
                                f"&bsns_year={year}&reprt_code=11011&fs_div={fs}").decode("utf-8", "replace"))
        except Exception:
            continue
        if d.get("status") != "000" or not d.get("list"):
            continue
        cur, prv = {}, {}
        for it in d["list"]:
            nm = (it.get("account_nm") or "").replace(" ", "")
            aid = it.get("account_id") or ""
            c, p = _num(it.get("thstrm_amount")), _num(it.get("frmtrm_amount"))
            if c is None:
                continue
            key_ = None
            if nm in ("매출액", "수익(매출액)", "영업수익") or aid == "ifrs-full_Revenue":
                key_ = "revenue"
            elif nm == "영업이익" or aid in ("dart_OperatingIncomeLoss",
                                          "ifrs-full_ProfitLossFromOperatingActivities"):
                key_ = "operating_income"
            elif nm in ("당기순이익", "당기순이익(손실)") or aid == "ifrs-full_ProfitLoss":
                key_ = "net_income"
            if key_ and key_ not in cur:
                cur[key_] = c
                if p is not None:
                    prv[key_] = p
        if cur.get("revenue"):
            out = {"fs_div": fs, **cur}
            for k in ("revenue", "operating_income", "net_income"):
                if cur.get(k) is not None and prv.get(k):
                    out[f"{k}_prev"] = prv[k]
                    out[f"{k}_yoy_pct"] = round((cur[k] / prv[k] - 1) * 100, 1) if prv[k] else None
            return out
    return None


def build(code, name, key, cmap):
    corp = cmap.get(code)
    if not corp:
        print(f"  [WARN] {name}({code}) 고유번호 없음")
        return None
    now_y = datetime.now(KST).year
    for y in range(now_y - 1, now_y - 1 - YEARS, -1):
        emp = employees(key, corp, y)
        fin = financials(key, corp, y)
        if not emp or not fin:
            continue
        rev, op = fin["revenue"], fin.get("operating_income")
        pay = emp.get("payroll_total")
        # 노동생산성 대용치 — 1인당 매출 증가율. 직원 수 전년치가 없으면 매출 증가율로 근사.
        prod = fin.get("revenue_yoy_pct")
        prev_emp = employees(key, corp, y - 1)
        if prev_emp and prev_emp.get("headcount") and fin.get("revenue_prev"):
            rph_now = rev / emp["headcount"]
            rph_prv = fin["revenue_prev"] / prev_emp["headcount"]
            prod = round((rph_now / rph_prv - 1) * 100, 1) if rph_prv else prod
        out = {
            "code": code, "name": name, "year": y, "fs_div": fin["fs_div"],
            "headcount": emp["headcount"], "avg_pay": emp.get("avg_pay"),
            "headcount_prev": (prev_emp or {}).get("headcount"),
            "regular": emp.get("regular"), "temporary": emp.get("temporary"),
            "temp_ratio_pct": emp.get("temp_ratio_pct"), "avg_tenure_yr": emp.get("avg_tenure_yr"),
            "payroll_total": pay, "revenue": rev, "operating_income": op,
            "net_income": fin.get("net_income"),
            "revenue_yoy_pct": fin.get("revenue_yoy_pct"),
            "op_yoy_pct": fin.get("operating_income_yoy_pct"),
            "net_yoy_pct": fin.get("net_income_yoy_pct"),
            "productivity_yoy_pct": prod,          # 1인당 매출 증가율 (노동생산성 대용)
            "payroll_to_revenue_pct": round(pay / rev * 100, 2) if (pay and rev) else None,
            "op_margin_pct": round(op / rev * 100, 2) if (op and rev) else None,
            "net_margin_pct": round(fin["net_income"] / rev * 100, 2) if (fin.get("net_income") and rev) else None,
            "revenue_per_head": round(rev / emp["headcount"]) if emp["headcount"] else None,
        }
        # 인상률 시나리오 — 인건비 증가가 영업이익률을 얼마나 깎는지
        if pay and rev and op is not None:
            out["scenarios"] = [
                {"raise_pct": r,
                 "payroll_add": round(pay * r / 100),
                 "op_after": round(op - pay * r / 100),
                 "op_margin_after_pct": round((op - pay * r / 100) / rev * 100, 2),
                 "margin_drop_pp": round(pay * r / 100 / rev * 100, 2)}
                for r in (2, 3, 5, 7, 10)]
        print(f"  {name}: {y}년 직원 {emp['headcount']:,}명 · 1인평균 "
              f"{(emp.get('avg_pay') or 0)/1e6:.1f}백만원 · 인건비/매출 "
              f"{out['payroll_to_revenue_pct']}% · 영업이익률 {out['op_margin_pct']}%")
        print(f"      전년비 매출 {out['revenue_yoy_pct']}% · 영업이익 {out['op_yoy_pct']}% · "
              f"순이익 {out['net_yoy_pct']}% · 1인당매출 {out['productivity_yoy_pct']}%")
        if out.get("temp_ratio_pct") is not None:
            print(f"      고용형태 정규직 {out['regular']:,}명 · 기간제 {out['temporary']:,}명 "
                  f"(비정규 비율 {out['temp_ratio_pct']}%) · 평균근속 {out['avg_tenure_yr']}년")
        return out
    print(f"  [WARN] {name} 최근 {YEARS}개 연도 데이터 없음")
    return None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    now = datetime.now(KST)
    print("=" * 60)
    print("  회사별 임금협상 근거자료 (DART 직원현황 + 손익)")
    print("=" * 60)

    key = _key()
    if not key:
        print("[INFO] DART_API_KEY 없음 — 수집 생략(로컬). GitHub Actions에서는 수집됨.")
        return 0
    try:
        cmap = corp_map(key)
        print(f"  기업 고유번호 {len(cmap):,}건 로드")
    except Exception as e:
        print(f"[ERROR] corpCode 실패: {e}")
        return 1

    rows = []
    for code, name in WATCH:
        try:
            r = build(code, name, key, cmap)
            if r:
                rows.append(r)
        except Exception as e:
            print(f"  [WARN] {name} 실패: {str(e)[:60]}")
    if not rows:
        print("[ERROR] 전 종목 실패 — 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "companies": rows,
        "method": {
            "source": "DART 사업보고서(11011) — empSttus(직원현황) + fnlttSinglAcntAll(손익)",
            "payroll_to_revenue": "연간급여총액 ÷ 매출액 — 인상 여력을 보는 핵심 비율",
            "scenario": "인상률 X% 적용 시 인건비 증가액과 영업이익률 하락폭(%p)",
        },
        "caveats": [
            "상장사만 공시 의무가 있어 비상장은 조회되지 않는다 — 화면에서 직접 입력해야 한다.",
            "공시 급여총액의 포함 범위(등기임원 보수·상여·복리후생 등)가 회사마다 달라 단순 비교에 주의가 필요하다.",
            "매출은 연결(CFS) 우선인데 직원현황은 별도 기준일 수 있어 분모·분자 범위가 어긋날 수 있다.",
            "시나리오는 '다른 조건이 그대로일 때'의 산술 계산이며, 매출 변동·생산성 변화는 반영하지 않는다.",
        ],
        "neutrality": ("생계비(물가)와 지불능력(재무)을 같은 화면에 놓기 위한 자료다. "
                       "어느 쪽 주장도 대변하지 않으며, 적정 임금은 생산성·업계 관행 등 "
                       "여기 없는 요소도 함께 따져야 한다."),
        "note": "공개 공시자료의 규칙기반 계산 · 노무·법률 자문 아님. 제출 전 원문 사업보고서로 교차 확인 필요.",
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"\n[OK] {OUTPUT_FILE} ({len(rows)}개사)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
