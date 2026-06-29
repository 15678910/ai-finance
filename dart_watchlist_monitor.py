"""
관심종목 DART 수주·투자 공시 모니터 — 신규 공급계약/시설투자 텔레그램 경보
=============================================================================
대통령 3대 메가프로젝트 수혜 관심종목(watchlist)에 한해 DART 전자공시를 주기적으로 조회.
'단일판매·공급계약', '신규 시설투자', '수주', '증설' 등 정책 실체를 확인하는 공시가 뜨면
신규분만 텔레그램으로 즉시 알림(중복 방지). 발표(기대)가 아니라 '실제 계약·투자'를 추적한다.

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
        if data.get("status") == "000":
            return data.get("list", [])
    except Exception as e:
        print(f"    [공시 조회 실패] {corp_code}: {e}")
    return []


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
    print("  관심종목 DART 수주·투자 공시 모니터")
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
    for sc in codes:
        cc = corp_map.get(sc)
        if not cc:
            continue
        corp_code, corp_name = cc
        for d in list_disclosures(api_key, corp_code, bgn, end):
            rnm = (d.get("report_nm") or "").strip()
            if not KW.search(rnm):
                continue
            rcept = (d.get("rcept_no") or "").strip()
            found.append({
                "stock_code": sc, "name": name_of.get(sc, corp_name),
                "report_nm": rnm, "rcept_dt": (d.get("rcept_dt") or "").strip(),
                "rcept_no": rcept, "flr_nm": (d.get("flr_nm") or "").strip(),
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
    seen = set(st.get("seen", []))
    for f in found:
        if f["rcept_no"] and f["rcept_no"] not in seen:
            new_items.append(f)
            seen.add(f["rcept_no"])

    print(f"  공시 {len(found)}건(키워드 필터) · 신규 {len(new_items)}건")
    for f in found[:12]:
        flag = "🆕" if f in new_items else "  "
        print(f"  {flag} [{f['rcept_dt']}] {f['name']}: {f['report_nm']}")

    # 텔레그램 경보(신규 — 최대 8건)
    if new_items:
        try:
            from core import send_message, get_secret
            if get_secret("TELEGRAM_FINANCE_BOT_TOKEN"):
                lines = [f"• [{f['name']}] {f['report_nm']} ({f['rcept_dt']})\n  {f['url']}"
                         for f in new_items[:8]]
                body = "📢 관심종목 신규 공시 (DART 수주·투자)\n" + "\n".join(lines)
                if len(new_items) > 8:
                    body += f"\n…외 {len(new_items)-8}건"
                if send_message(body):
                    print(f"  📨 신규 공시 경보 {len(new_items)}건 발송")
        except Exception as e:
            print(f"  [WARN] 텔레그램 실패: {e}")

    # 상태 저장(최근 600건 한정)
    try:
        from core import save_state
        save_state("dart_watchlist", {"seen": list(seen)[-600:]})
    except Exception:
        pass

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "n": len(found), "disclosures": found[:MAX_KEEP],
        "note": ("DART 전자공시(opendart.fss.or.kr) — 관심종목의 단일판매·공급계약/시설투자/수주 등 "
                 "정책 실체 확인용 공시. '발표(기대)'가 아니라 '실제 계약·투자'를 추적. "
                 "원문 확인 필수·투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
