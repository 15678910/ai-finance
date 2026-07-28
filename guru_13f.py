"""
거장 매매 추적 — SEC 13F 보유 공시 (Valley '거장 매매' 대응)
=============================================================
SEC EDGAR 무료 API로 유명 투자자(버핏·드러켄밀러·버리·달리오·애크먼)의
분기 13F-HR 보유 내역을 수집, 직전 분기와 비교해 신규/확대/축소/청산을 판별.

  · 13F는 분기 종료 후 최대 45일 지연 공시 — '실시간' 아님
  · 미국 상장 롱 포지션만 (공매도·해외주식·채권 미포함)
  · value 단위: USD (2023년 이후 원단위 보고)

출력: docs/guru_13f.json  (GitHub Actions 실행 가능 — SEC는 IP 차단 없음, UA 필수)
🚨 정보 제공용 · 투자자문 아님.
"""

import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "guru_13f.json")

# (표시명, 기관명, CIK)
GURUS = [
    ("워런 버핏", "Berkshire Hathaway", "1067983"),
    ("스탠리 드러켄밀러", "Duquesne Family Office", "1536411"),
    ("마이클 버리", "Scion Asset Management", "1649339"),
    ("레이 달리오 (창업)", "Bridgewater Associates", "1350694"),
    ("빌 애크먼", "Pershing Square", "1336528"),
]
UA = {"User-Agent": "ai-finance-dashboard research lacoiffure828@gmail.com"}


def _get(url, timeout=20):
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, timeout=timeout).read()


def _get_json(url):
    return json.loads(_get(url).decode("utf-8", "replace"))


def latest_13f_accessions(cik):
    """최근 13F-HR 2개 분기 [(reportDate, accession)] — 같은 분기 정정(/A)은 최신 우선."""
    subs = _get_json(f"https://data.sec.gov/submissions/CIK{int(cik):010d}.json")
    rec = subs.get("filings", {}).get("recent", {})
    forms = rec.get("form", [])
    accs = rec.get("accessionNumber", [])
    rdates = rec.get("reportDate", [])
    by_period = {}                                  # reportDate -> accession (최신 파일 우선, recent는 최신순)
    for i, fm in enumerate(forms):
        if fm in ("13F-HR", "13F-HR/A"):
            rd = rdates[i]
            if rd not in by_period:
                by_period[rd] = accs[i]
    periods = sorted(by_period.keys(), reverse=True)[:2]
    return [(p, by_period[p]) for p in periods]


def fetch_holdings(cik, accession):
    """13F infotable XML → {issuer: {value_usd, shares}}"""
    acc = accession.replace("-", "")
    base = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc}"
    idx = _get_json(f"{base}/index.json")
    xml_files = [it["name"] for it in idx.get("directory", {}).get("item", [])
                 if it["name"].lower().endswith(".xml") and "primary_doc" not in it["name"].lower()]
    if not xml_files:
        return None
    # 보통 infotable.xml / form13fInfoTable.xml — 첫 번째 시도 후 infoTable 태그 없으면 다음
    for fn in xml_files:
        raw = _get(f"{base}/{fn}").decode("utf-8", "replace")
        if "infoTable" not in raw:
            continue
        holds = {}
        for m in re.finditer(r"<(?:\w+:)?infoTable>(.*?)</(?:\w+:)?infoTable>", raw, re.S):
            blk = m.group(1)
            nm = re.search(r"<(?:\w+:)?nameOfIssuer>(.*?)</", blk, re.S)
            vl = re.search(r"<(?:\w+:)?value>([\d.]+)</", blk)
            sh = re.search(r"<(?:\w+:)?sshPrnamt>([\d.]+)</", blk)
            put_call = re.search(r"<(?:\w+:)?putCall>(.*?)</", blk)
            if not nm or not vl:
                continue
            name = re.sub(r"\s+", " ", nm.group(1)).strip().title()
            if put_call:                            # 풋/콜 옵션은 별도 표기
                name += f" ({put_call.group(1).strip().upper()})"
            d = holds.setdefault(name, {"value": 0.0, "shares": 0.0})
            d["value"] += float(vl.group(1))
            d["shares"] += float(sh.group(1)) if sh else 0.0
        if holds:
            return holds
    return None


def diff_holdings(cur, prev):
    """신규/청산/확대/축소 (주식수 기준, ±15% 임계)."""
    new, exited, added, reduced = [], [], [], []
    for n, d in cur.items():
        p = prev.get(n)
        if p is None:
            new.append((n, d["value"]))
        elif p["shares"] > 0 and d["shares"] > p["shares"] * 1.15:
            added.append((n, d["value"], (d["shares"] / p["shares"] - 1) * 100))
        elif p["shares"] > 0 and d["shares"] < p["shares"] * 0.85:
            reduced.append((n, d["value"], (d["shares"] / p["shares"] - 1) * 100))
    for n, p in prev.items():
        if n not in cur:
            exited.append((n, p["value"]))
    new.sort(key=lambda x: -x[1])
    exited.sort(key=lambda x: -x[1])
    added.sort(key=lambda x: -x[1])
    reduced.sort(key=lambda x: -x[1])
    return new, exited, added, reduced


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    gurus_out = []
    for disp, org, cik in GURUS:
        try:
            pairs = latest_13f_accessions(cik)
            if not pairs:
                print(f"  [WARN] {org}: 13F 없음")
                continue
            cur_p, cur_a = pairs[0]
            cur = fetch_holdings(cik, cur_a)
            time.sleep(0.4)                          # SEC rate limit 예의
            prev, prev_p = None, None
            if len(pairs) > 1:
                prev_p, prev_a = pairs[1]
                prev = fetch_holdings(cik, prev_a)
                time.sleep(0.4)
            if not cur:
                print(f"  [WARN] {org}: infotable 파싱 실패")
                continue

            total = sum(d["value"] for d in cur.values()) or 1
            # 일부 파일은 구형식(천 달러 단위)으로 보고 → 총액 $1억 미만이면 ×1000 보정
            if total < 1e8:
                for d in cur.values():
                    d["value"] *= 1000
                if prev:
                    for d in prev.values():
                        d["value"] *= 1000
                total *= 1000
            top = sorted(cur.items(), key=lambda kv: -kv[1]["value"])[:10]
            top_out = [{"name": n, "value_usd": round(d["value"]),
                        "pct": round(d["value"] / total * 100, 1),
                        "shares": round(d["shares"])} for n, d in top]
            changes = {}
            if prev:
                new, exited, added, reduced = diff_holdings(cur, prev)
                changes = {
                    "new": [{"name": n, "value_usd": round(v)} for n, v in new[:6]],
                    "exited": [{"name": n, "value_usd": round(v)} for n, v in exited[:6]],
                    "added": [{"name": n, "value_usd": round(v), "chg_pct": round(c)} for n, v, c in added[:6]],
                    "reduced": [{"name": n, "value_usd": round(v), "chg_pct": round(c)} for n, v, c in reduced[:6]],
                }
            gurus_out.append({
                "guru": disp, "org": org, "cik": cik,
                "period": cur_p, "prev_period": prev_p,
                "total_value_usd": round(total),
                "n_positions": len(cur),
                "top": top_out, "changes": changes,
            })
            print(f"  {disp} ({org}): {cur_p} 기준 {len(cur)}종목 · ${total/1e9:.1f}B · "
                  f"신규 {len(changes.get('new', []))} 청산 {len(changes.get('exited', []))}")
        except Exception as e:
            print(f"  [WARN] {org} 실패: {type(e).__name__} {str(e)[:80]}")
            continue

    if not gurus_out:
        print("[ERROR] 전체 실패 — 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "gurus": gurus_out,
        "note": ("SEC EDGAR 13F-HR 분기 보유 공시. 분기 종료 후 최대 45일 지연 — 실시간 아님. "
                 "미국 상장 롱 포지션만(공매도·해외·채권 제외). 신규/확대/축소/청산은 직전 분기 주식수 대비(±15%). "
                 "(PUT)/(CALL)=옵션 포지션. 정보 제공용 · 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE} ({len(gurus_out)}명)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
