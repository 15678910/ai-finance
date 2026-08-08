"""
코스피 밸류에이션 자동 수집 — 후행 PER · PBR · 배당수익률 (+ 가능 시 선행 PER)
================================================================================
1순위) KRX 정보데이터시스템 (pykrx 로그인) — 원본·정확. KRX가 2024년경부터 데이터
       API에 회원 로그인을 요구 → GitHub Secrets의 KRX_ID/KRX_PW 로 인증.
       get_index_fundamental("1001"=코스피) → PER(후행)·선행PER·PBR·배당수익률.
       ※ 로그인 방식이라 IP 차단과 무관 → GitHub Actions에서도 동작 기대.
2순위) INDEXerGO 스크래핑 (원출처 KRX·KOFIA) — KRX 로그인 없을 때/실패 시 폴백.
       단 GitHub 데이터센터 IP는 지역차단(403)되므로 로컬(한국 IP)에서만 유효.

둘 다 실패하면 기존 파일을 보존(과거값 유지).

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
KOSPI_TICKER = "1001"                       # KRX 코스피 지수 티커
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120"


# ────────────────────────────── 1순위: KRX (pykrx 로그인) ──────────────────────────────
def from_krx():
    """KRX 로그인 기반 코스피 지수 밸류에이션. (dict) 또는 실패 시 None."""
    if not (os.environ.get("KRX_ID") and os.environ.get("KRX_PW")):
        print("  [INFO] KRX_ID/KRX_PW 미설정 — KRX 자동수집 건너뜀(폴백 시도).")
        return None
    try:
        from pykrx import stock
    except Exception as e:
        print(f"  [WARN] pykrx import 실패: {e}")
        return None

    now = datetime.now(KST)
    frm = (now - timedelta(days=12)).strftime("%Y%m%d")   # 최근 영업일 확보용 여유
    to = now.strftime("%Y%m%d")
    try:
        df = stock.get_index_fundamental(frm, to, KOSPI_TICKER)
    except Exception as e:
        print(f"  [WARN] KRX 조회 실패(로그인/네트워크): {e}")
        return None
    if df is None or len(df) == 0:
        print("  [WARN] KRX 응답 비어있음 — 로그인 실패 또는 데이터 없음.")
        return None

    row = df.iloc[-1]                                     # 최신 영업일
    idx = df.index[-1]
    asof = idx.strftime("%Y-%m-%d") if hasattr(idx, "strftime") else str(idx)[:10]

    def g(col):
        """KRX 값 → 소수 2자리. pykrx가 float32로 주기 때문에 그대로 float()하면
        18.03이 18.030000686645508처럼 이진 오차가 그대로 노출된다(KRX 원본은 2자리)."""
        try:
            v = float(row[col])
            return round(v, 2) if v > 0 else None
        except Exception:
            return None

    per, pbr, div, fwd = g("PER"), g("PBR"), g("배당수익률"), g("선행PER")
    if per is None and pbr is None:
        print("  [WARN] KRX PER/PBR 파싱 실패.")
        return None

    out = {
        "per": per, "pbr": pbr, "div_yield": div, "asof": asof,
        "source": "KRX 정보데이터시스템 (pykrx 로그인)",
        "source_url": "https://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd?menuId=MDC0201020506",
    }
    if fwd is not None:                                   # KRX가 선행PER 제공 시(종종 '-'라 없을 수 있음)
        out["fwd_per"] = fwd
    print(f"  [KRX] 후행PER {per} · 선행PER {fwd if fwd is not None else '—'} · PBR {pbr} · 배당 {div} ({asof})")
    return out


# ────────────────────────────── 2순위: INDEXerGO 폴백 ──────────────────────────────
URL = "https://www.indexergo.com/series/?frq=D&idxDetail={}"
ITEMS = [("per", "후행 PER", 20205), ("pbr", "PBR", 20206), ("div_yield", "배당수익률", 20207)]
PAT = re.compile(r"(\d{4})\.(\d{2})\.(\d{2})\s*마감 기준\s*[^:]{1,12}:\s*([\d.]+)")
HEADERS = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://www.indexergo.com/",
}


def _scrape_igo(idx):
    try:
        req = urllib.request.Request(URL.format(idx), headers=HEADERS)
        html = urllib.request.urlopen(req, timeout=15).read().decode("utf-8", errors="replace")
        t = re.sub(r"<script.*?</script>", "", html, flags=re.S)
        t = re.sub(r"<[^>]+>", " ", t)
        t = re.sub(r"\s+", " ", t)
        if "| KOSPI |" not in t:
            return None, None
        m = PAT.search(t)
        if not m:
            return None, None
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}", float(m.group(4))
    except Exception as e:
        print(f"  [WARN] INDEXerGO idxDetail={idx} 실패: {e}")
        return None, None


def from_indexergo():
    out, asof_seen, got = {}, [], 0
    for key, label, idx in ITEMS:
        asof, val = _scrape_igo(idx)
        out[key] = val
        if val is not None:
            got += 1
            asof_seen.append(asof)
            print(f"  [IGO] 코스피 {label}: {val} ({asof})")
    if got == 0:
        return None
    out["asof"] = max(asof_seen) if asof_seen else None
    out["source"] = "INDEXerGO (원출처: KRX·KOFIA)"
    out["source_url"] = "https://www.indexergo.com/series/?frq=D&idxDetail=20205"
    return out


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    data = from_krx()                        # 1순위
    if data is None:
        print("  → INDEXerGO 폴백 시도")
        data = from_indexergo()              # 2순위
    if data is None:
        print("[ERROR] KRX·INDEXerGO 모두 실패 — 기존 파일 보존.")
        return 1

    # ── 후퇴 방지 ────────────────────────────────────────────────────
    # 폴백(INDEXerGO)은 원본보다 하루 늦게 갱신된다. KRX 파싱이 한 번 실패하면
    # 이미 갖고 있던 최신 자료를 '더 낡은 값'으로 덮어쓰게 된다.
    # 데이터가 채워지긴 하므로 정체 경보에도 안 걸려 조용히 하루 뒤처진다.
    # (2026-08-07 실제 발생: KRX 파싱 실패 → 08-06 자료로 후퇴)
    prev = {}
    try:
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            prev = json.load(f)
    except Exception:
        pass
    if prev.get("asof") and data.get("asof") and data["asof"] < prev["asof"]:
        print(f"  [WARN] 새 자료({data['asof']})가 기존({prev['asof']})보다 오래됨 — "
              f"덮어쓰지 않고 기존 파일 보존. (출처: {data.get('source')})")
        return 0

    # 직전 거래일보다 뒤처지면 화면에서 알 수 있게 표시한다.
    data["asof_lag_days"] = None
    try:
        from core.expiry import is_closed
        d = now.date()
        for _ in range(10):                       # 직전 거래일 탐색
            d -= timedelta(days=1)
            if not is_closed(d, "KRX"):
                break
        if data.get("asof"):
            gap = (d - datetime.strptime(data["asof"], "%Y-%m-%d").date()).days
            data["asof_lag_days"] = max(0, gap)
            if gap > 0:
                print(f"  [WARN] 기준일 {data['asof']} 는 직전 거래일({d})보다 {gap}일 뒤짐")
    except Exception:
        pass

    data["generated_at"] = now.strftime("%Y-%m-%d %H:%M:%S KST")
    data["note"] = (
        "코스피 지수 후행 PER·PBR·배당수익률 (마감 기준, 일 1회). "
        "1순위 KRX 정보데이터시스템(로그인), 실패 시 INDEXerGO(KRX·KOFIA) 폴백. "
        "선행(Forward) PER은 KRX가 제공 시 fwd_per로 포함(종종 미제공) — 없으면 대시보드에서 수동 입력. "
        "실시간 아님 · 투자자문 아님."
    )

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE}  (source={data.get('source')} · 기준 {data.get('asof')})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
