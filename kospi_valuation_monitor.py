"""
코스피 밸류에이션 자동 수집 — 후행 PER · PBR · 배당수익률
==========================================================
KRX 정보데이터시스템은 nProtect 보안프로그램을 요구해 자동수집 불가(사람도 접속 곤란).
→ INDEXerGO(원출처: KRX·KOFIA)에서 코스피 지수 밸류에이션을 스크래핑.

  idxDetail=20205 → 코스피 후행 PER
  idxDetail=20206 → 코스피 PBR
  idxDetail=20207 → 코스피 배당수익률(%)

※ 선행(Forward) PER은 증권사 컨센서스(FnGuide) 기반이라 무료 공개원이 없음 → 대시보드에서 수동 입력.

출력: docs/kospi_valuation.json
🚨 참고용(일 1회 마감 기준, 실시간 아님). 투자자문 아님.
"""

import json
import os
import re
import sys
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "kospi_valuation.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120"
URL = "https://www.indexergo.com/series/?frq=D&idxDetail={}"

# (JSON 키, 라벨, idxDetail)
ITEMS = [("per", "후행 PER", 20205), ("pbr", "PBR", 20206), ("div_yield", "배당수익률", 20207)]

# "2026.07.14 마감 기준 PER: 18.65" + "| KOSPI |" 로 시장 확인
PAT = re.compile(r"(\d{4})\.(\d{2})\.(\d{2})\s*마감 기준\s*[^:]{1,12}:\s*([\d.]+)")


HEADERS = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://www.indexergo.com/",
    "Sec-Ch-Ua": '"Chromium";v="120", "Not(A:Brand";v="24", "Google Chrome";v="120"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "document", "Sec-Fetch-Mode": "navigate", "Sec-Fetch-Site": "same-origin",
    "Upgrade-Insecure-Requests": "1",
}


def scrape(idx):
    """(asof 'YYYY-MM-DD', value float) — 실패 시 (None, None)."""
    try:
        req = urllib.request.Request(URL.format(idx), headers=HEADERS)
        html = urllib.request.urlopen(req, timeout=15).read().decode("utf-8", errors="replace")
        t = re.sub(r"<script.*?</script>", "", html, flags=re.S)
        t = re.sub(r"<[^>]+>", " ", t)
        t = re.sub(r"\s+", " ", t)
        if "| KOSPI |" not in t:                       # 코스피 시리즈가 맞는지 확인
            print(f"  [WARN] idxDetail={idx}: KOSPI 시리즈 아님 — 스킵")
            return None, None
        m = PAT.search(t)
        if not m:
            print(f"  [WARN] idxDetail={idx}: 값 패턴 불일치(사이트 구조 변경 의심)")
            return None, None
        asof = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        return asof, float(m.group(4))
    except Exception as e:
        print(f"  [WARN] idxDetail={idx} 실패: {e}")
        return None, None


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    out = {"generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST")}
    asof_seen, got = [], 0
    for key, label, idx in ITEMS:
        asof, val = scrape(idx)
        out[key] = val
        if val is not None:
            got += 1
            asof_seen.append(asof)
            print(f"  코스피 {label}: {val} ({asof})")

    if got == 0:
        print("[ERROR] 전부 수집 실패 — 기존 파일 보존.")
        return 1

    out["asof"] = max(asof_seen) if asof_seen else None
    out["source"] = "INDEXerGO (원출처: KRX·KOFIA)"
    out["source_url"] = "https://www.indexergo.com/series/?frq=D&idxDetail=20205"
    out["note"] = ("코스피 지수 후행 PER·PBR·배당수익률 (마감 기준, 일 1회). "
                   "선행(Forward) PER은 증권사 컨센서스라 무료 공개원이 없어 대시보드에서 수동 입력. "
                   "실시간 아님 · 투자자문 아님.")

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE}  ({got}/{len(ITEMS)}개 · 기준 {out['asof']})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
