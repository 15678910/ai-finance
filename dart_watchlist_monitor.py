"""
관심종목 DART 공시 모니터 — 수주·투자(기회)와 악재(위험)를 함께 추적
=============================================================================
대통령 3대 메가프로젝트 수혜 관심종목(watchlist)에 한해 DART 전자공시를 주기적으로 조회.
두 방향을 모두 본다. 한쪽만 보면 판단이 기운다.
  · 기회 — '단일판매·공급계약', '신규 시설투자', '수주', '증설' 등 정책 실체를 확인하는 공시.
           발표(기대)가 아니라 '실제 계약·투자'를 추적한다.
  · 위험 — 유상증자·전환사채(지분 희석), 감사의견 거절·자본잠식(존속), 횡령·배임(위법),
           불성실공시·과징금(제재) 등. 수주 공시만 모으면 악재가 구조적으로 안 보인다.
신규분만 텔레그램으로 알리고(중복 방지) 악재는 별도 메시지로 분리해 보낸다.

DART OpenAPI(무료, opendart.fss.or.kr) — DART_API_KEY 필요.
출력: docs/dart_disclosures.json
🚨 정보 모니터링용. 공시 원문 확인 필수. 투자 결정 단독 사용 금지.
"""

import json
import os
import re
import sys
import zipfile
import urllib.request
import xml.etree.ElementTree as ET
from io import BytesIO
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "dart_disclosures.json")
DART_BASE = "https://opendart.fss.or.kr/api"
DAYS_BACK = 4          # 주말 공백 커버
MAX_KEEP = 40          # JSON 보관 건수

# 수주·투자 관련 공시명 필터(정책 실체 확인용)
KW = re.compile(r"공급계약|단일판매|수주|시설투자|증설|유형자산\s*취득|투자판단|타법인주식|영업양수|착공|투자협약|MOU|신규시설")

# 악재성 공시 — 등급 기준은 '주주가치 손상이 얼마나 되돌리기 어려운가'다.
#   high  : 존속·상장 자체가 걸린 것, 또는 경영진의 위법이 드러난 것
#   mid   : 지분이 실제로 희석되거나 제재가 확정된 것
#   low   : 결과가 아직 열려 있어 관찰만 필요한 것
# 순서가 곧 우선순위다(위에서 먼저 맞으면 그 등급으로 확정).
# (등급, 구분, 패턴, 제외 패턴, 해소 판정 적용 여부)
#   제외 패턴 — 같은 단어를 쓰지만 악재가 아닌 일상 공시(예: '전환사채 만기전 사채취득'은
#               희석이 아니라 오히려 희석 감소).
#   해소 판정 — '계약해지'의 '해지/해제'는 그 자체가 악재(수주 취소)라 해소로 보면 안 된다.
# '대표이사 변경'·'채무보증'은 그룹사에서 일상적으로 나와 경보 피로만 키워서 넣지 않았다.
BAD_RULES = [
    ("high", "존속·상장", r"의견\s*거절|한정의견|감사의견\s*(?:비적정|거절|한정)|계속기업|존속\s*능력"
                         r"|자본\s*잠식|상장\s*폐지|상장폐지|거래\s*정지|매매거래\s*정지"
                         r"|관리종목|투자주의\s*환기|회생\s*절차|파산|워크아웃|부도|채권은행", None, True),
    ("high", "위법·부정", r"횡령|배임|분식|부정\s*거래|시세\s*조종|주가\s*조작|구속(?!력)", None, True),
    ("mid",  "수주 취소", r"공급\s*계약\s*(?:해지|해제)|단일판매.*(?:해지|해제)", None, False),
    ("mid",  "지분 희석", r"유상증자|전환사채|신주인수권부사채|교환사채|무상감자|유상감자|자본\s*감소",
                         r"사채\s*취득|만기\s*전|전환가액\s*의?\s*조정|행사가액\s*의?\s*조정", True),
    ("mid",  "제재·행정", r"불성실공시|공시위반|과징금|과태료|시정\s*명령|영업\s*정지|행정\s*처분"
                         r"|검찰|압수|기소|고발|세무조사", None, True),
    ("low",  "분쟁·변동", r"소송|가처분|조회\s*공시|최대주주\s*변경|주식\s*담보|타인에\s*대한\s*담보"
                         r"|리콜|영업\s*양도", None, True),
]
BAD_RULES = [(sev, cat, re.compile(p), re.compile(x) if x else None, res)
             for sev, cat, p, x, res in BAD_RULES]

# 악재 '해소' 공시는 같은 단어를 쓴다(예: '거래정지 해제', '유상증자 철회', '가처분 기각').
# 이걸 걸러내지 않으면 호재를 악재로 알린다.
RESOLVED = re.compile(r"해제|해소|취소|철회|기각|각하|승소|무혐의|불기소|종결|면제")
# 단, 회사가 낸 불복이 기각된 것은 악재가 확정된 것이다('상장폐지 이의신청 기각').
APPEAL_LOST = re.compile(r"(?:이의\s*신청|항고|항소|상고|재심|불복).*(?:기각|각하)")

SEV_ICON = {"high": "🔴", "mid": "🟠", "low": "🟡"}
SEV_LABEL = {"high": "높음", "mid": "중간", "low": "낮음"}
SEV_ORDER = {"high": 0, "mid": 1, "low": 2}


def classify(report_nm):
    """공시명 → (kind, severity, category). 해당 없으면 (None, None, None).

    악재를 먼저 본다. 한 공시가 양쪽 키워드를 다 가질 때(예: '유상증자 결정'으로 조달한
    '시설투자') 투자자에게 먼저 필요한 정보는 희석 쪽이기 때문이다.
    """
    resolved_cat = None
    for sev, cat, pat, exclude, resolvable in BAD_RULES:
        if not pat.search(report_nm) or (exclude and exclude.search(report_nm)):
            continue
        if (resolvable and RESOLVED.search(report_nm)
                and not APPEAL_LOST.search(report_nm)):
            # 이 규칙 기준으로는 해소다. 다른 규칙이 진짜 악재로 잡는지 끝까지 본다.
            resolved_cat = resolved_cat or cat
            continue
        return ("악재", sev, cat)
    if resolved_cat:
        return ("악재해소", None, resolved_cat)
    if KW.search(report_nm):
        return ("수주·투자", None, None)
    return (None, None, None)


def _get_key():
    k = os.environ.get("DART_API_KEY")
    if k:
        return k.strip()
    try:
        from core import get_secret
        return (get_secret("DART_API_KEY") or "").strip() or None
    except Exception:
        return None


def get_corp_map(api_key, want_codes):
    """corpCode.xml(전체) → {stock_code: (corp_code, corp_name)} (관심종목만)."""
    url = f"{DART_BASE}/corpCode.xml?crtfc_key={api_key}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=40) as resp:
        content = resp.read()
    with zipfile.ZipFile(BytesIO(content)) as zf:
        with zf.open(zf.namelist()[0]) as xf:
            root = ET.parse(xf).getroot()
    want = set(want_codes)
    out = {}
    for el in root.iter("list"):
        sc = (el.findtext("stock_code") or "").strip()
        if sc in want:
            out[sc] = ((el.findtext("corp_code") or "").strip(),
                       (el.findtext("corp_name") or "").strip())
    return out


def list_disclosures(api_key, corp_code, bgn, end):
    url = (f"{DART_BASE}/list.json?crtfc_key={api_key}"
           f"&corp_code={corp_code}&bgn_de={bgn}&end_de={end}&page_count=50")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        status = data.get("status")
        if status == "000":
            return data.get("list", [])
        if status == "013":          # 조회된 데이터 없음 — 정상
            return []
        print(f"    [공시 조회 실패] {corp_code}: status={status} {data.get('message','')}")
    except Exception as e:
        print(f"    [공시 조회 실패] {corp_code}: {e}")
    # None = 조회 실패. 빈 목록([])과 구분해야 '악재 0건'을 거짓으로 띄우지 않는다.
    return None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    api_key = _get_key()
    if not api_key:
        print("[SKIP] DART_API_KEY 미설정 — 로컬에선 스킵(워크플로에서 실행). 기존 JSON 보존.")
        return 0

    try:
        from watchlist import WATCHLIST_STOCKS
    except Exception as e:
        print(f"[ERROR] watchlist 임포트 실패: {e}")
        return 1
    name_of = {c: n for n, c in WATCHLIST_STOCKS}
    codes = list(name_of.keys())

    print("=" * 55)
    print("  관심종목 DART 공시 모니터 (수주·투자 + 악재)")
    print(f"  KST {now.strftime('%Y-%m-%d %H:%M')}  · 종목 {len(codes)}")
    print("=" * 55)

    try:
        corp_map = get_corp_map(api_key, codes)
        print(f"  [DART] corp_code 매핑 {len(corp_map)}/{len(codes)}")
    except Exception as e:
        print(f"[ERROR] corpCode 다운로드 실패: {e}")
        return 1

    bgn = (now - timedelta(days=DAYS_BACK)).strftime("%Y%m%d")
    end = now.strftime("%Y%m%d")

    found = []
    failed = []   # 조회 실패 종목 — 있으면 '악재 0건'은 확인된 사실이 아니다
    for sc in codes:
        cc = corp_map.get(sc)
        if not cc:
            continue
        corp_code, corp_name = cc
        rows = list_disclosures(api_key, corp_code, bgn, end)
        if rows is None:
            failed.append(name_of.get(sc, corp_name))
            continue
        for d in rows:
            rnm = (d.get("report_nm") or "").strip()
            kind, sev, cat = classify(rnm)
            if not kind:
                continue
            rcept = (d.get("rcept_no") or "").strip()
            if not re.fullmatch(r"\d{14}", rcept):   # DART 접수번호=14자리 숫자. 변조·결측값이 URL/JSON에 들어가는 것 차단.
                continue
            found.append({
                "stock_code": sc, "name": name_of.get(sc, corp_name),
                "report_nm": rnm, "rcept_dt": (d.get("rcept_dt") or "").strip(),
                "rcept_no": rcept, "flr_nm": (d.get("flr_nm") or "").strip(),
                "kind": kind, "severity": sev, "category": cat,
                "url": f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept}",
            })

    # 최신순 정렬(접수일 desc, 접수번호 desc)
    found.sort(key=lambda x: (x["rcept_dt"], x["rcept_no"]), reverse=True)

    # 신규(미알림) 탐지
    new_items = []
    try:
        from core import load_state
        st = load_state("dart_watchlist", default={}) or {}
    except Exception:
        st = {}
    order = [r for r in (st.get("seen") or []) if isinstance(r, str)]
    seen = set(order)
    for f in found:
        if f["rcept_no"] and f["rcept_no"] not in seen:
            new_items.append(f)
    bad = [f for f in found if f["kind"] == "악재"]
    good = [f for f in found if f["kind"] == "수주·투자"]
    resolved = [f for f in found if f["kind"] == "악재해소"]
    new_bad = [f for f in new_items if f["kind"] == "악재"]
    new_good = [f for f in new_items if f["kind"] == "수주·투자"]
    # 악재는 등급순, 같은 등급 안에서는 최신순. found가 이미 최신순이고 sort는 안정 정렬이라
    # 등급 키 하나로만 다시 세우면 된다(날짜를 키에 넣으면 오래된 것이 앞으로 와서 잘린다).
    bad.sort(key=lambda x: SEV_ORDER.get(x["severity"], 9))
    new_bad.sort(key=lambda x: SEV_ORDER.get(x["severity"], 9))

    # seen 등록은 알림이 실제로 나간 뒤에만 한다. 기본값을 '전부 미발송'으로 두고 성공한 것만 지운다.
    # 반대로(실패한 것만 추가) 하면 토큰 미설정·예외 시 아무것도 추가되지 않아 전부 '알림 완료'로
    # 찍히고, 악재 경보가 조용히 사라진다. 메시지당 8건을 넘는 몫도 미발송으로 남아 다음 실행에 나간다.
    # 악재해소는 경보 대상이 아니므로 바로 seen 처리한다.
    unsent = {f["rcept_no"] for f in new_bad + new_good}

    print(f"  공시 {len(found)}건 (수주·투자 {len(good)} / 악재 {len(bad)} / 악재해소 {len(resolved)}) · 신규 {len(new_items)}건")
    for f in bad[:8]:
        flag = "🆕" if f in new_items else "  "
        print(f"  {flag} {SEV_ICON.get(f['severity'],'')} [{f['rcept_dt']}] {f['name']}: {f['report_nm']} ({f['category']})")
    for f in good[:8]:
        flag = "🆕" if f in new_items else "  "
        print(f"  {flag} 🟢 [{f['rcept_dt']}] {f['name']}: {f['report_nm']}")

    # 텔레그램 경보 — 악재와 수주를 분리해 보낸다. 섞어 보내면 악재가 묻힌다.
    # 각 발송을 따로 감싼다: 악재 발송이 예외를 던져도 수주 발송은 시도한다.
    def _send(label, items, header, fmt, footer=""):
        batch = items[:8]
        body = header + "\n" + "\n".join(fmt(f) for f in batch)
        if len(items) > 8:
            body += f"\n…외 {len(items)-8}건은 다음 실행에서 이어서 보냄"
        body += footer
        try:
            ok = send_message(body)
        except Exception as e:
            print(f"  [WARN] {label} 경보 예외: {e} — 다음 실행에서 재시도")
            return
        if ok:
            unsent.difference_update(f["rcept_no"] for f in batch)
            print(f"  📨 {label} 경보 {len(batch)}건 발송")
        else:
            print(f"  [WARN] {label} 경보 발송 실패 — 다음 실행에서 재시도")

    if unsent:
        try:
            from core import send_message, get_secret
            has_token = bool(get_secret("TELEGRAM_FINANCE_BOT_TOKEN"))
        except Exception as e:
            has_token = False
            print(f"  [WARN] 텔레그램 모듈 로드 실패: {e}")
        if not has_token:
            print(f"  [INFO] 텔레그램 토큰 없음 — 신규 {len(unsent)}건은 '미알림'으로 남김")
        else:
            if new_bad:
                _send("악재 공시", new_bad, "⚠️ 관심종목 악재 공시 (DART)",
                      lambda f: (f"{SEV_ICON.get(f['severity'],'')} [{f['name']}] {f['report_nm']}\n"
                                 f"  {f['category']}·위험 {SEV_LABEL.get(f['severity'],'')} "
                                 f"({f['rcept_dt']})\n  {f['url']}"),
                      "\n\n※ 공시 원문 확인 필수. 투자자문 아님.")
            if new_good:
                _send("수주·투자 공시", new_good, "📢 관심종목 신규 공시 (DART 수주·투자)",
                      lambda f: f"• [{f['name']}] {f['report_nm']} ({f['rcept_dt']})\n  {f['url']}")

    # 상태 저장(최근 600건 한정) — set이 아니라 리스트 순서를 유지해야 잘라낼 때
    # 오래된 것부터 빠진다. set을 그대로 자르면 최신 공시가 날아갈 수 있다.
    order += [f["rcept_no"] for f in new_items if f["rcept_no"] not in unsent]
    try:
        from core import save_state
        save_state("dart_watchlist", {"seen": order[-600:]})
    except Exception:
        pass

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "n": len(found), "days_back": DAYS_BACK,
        "n_bad": len(bad), "n_good": len(good),
        "n_failed": len(failed), "failed": failed[:30],
        # 악재는 수주 공시가 많은 날에도 밀려나지 않게 별도 배열로 보관한다.
        "bad": bad[:MAX_KEEP], "good": good[:MAX_KEEP], "resolved": resolved[:MAX_KEEP],
        "disclosures": found[:MAX_KEEP],
        "severity_legend": {"high": "존속·상장/위법 — 높음", "mid": "지분 희석·제재 — 중간",
                            "low": "분쟁·지배구조 변동 — 낮음"},
        "note": ("DART 전자공시(opendart.fss.or.kr) — 관심종목의 ①수주·투자(단일판매·공급계약/시설투자 등 "
                 "정책 실체 확인용)와 ②악재(유상증자·전환사채 등 지분 희석, 감사의견·자본잠식 등 존속, "
                 "횡령·배임 등 위법, 불성실공시·과징금 등 제재)를 함께 추적. 등급은 '손상의 되돌릴 수 없음' "
                 "기준의 기계적 분류이며 공시 원문 해석을 대신하지 않는다. "
                 "'해제/철회/기각' 등 악재 해소 공시는 경보에서 제외. 원문 확인 필수·투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
