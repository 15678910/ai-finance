"""
국내 증권사 리포트 수집기 — 네이버 금융 리서치 (증권사·목표주가·투자의견·날짜)
====================================================================================
종목별 최근 애널리스트 리포트를 주기적으로 수집 → 대시보드 + 신규/목표가변경 텔레그램 경보.
무료 소스(네이버 금융 리서치). 본문(PDF) 전문은 미제공 — 제목·증권사·목표가·투자의견까지.

출력: docs/analyst_reports.json
🚨 정보 모니터링용. 투자 결정 단독 사용 금지.
"""

import json
import os
import re
import sys
import time
import html as _html
import urllib.request
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "analyst_reports.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

# 관심종목 25개 전체(상장 24) — 대통령 3대 메가프로젝트 수혜 바스켓. SK·삼성 우선.
try:
    from watchlist import WATCHLIST_STOCKS
    _head = [("SK하이닉스", "000660"), ("삼성전자", "005930")]
    _rest = [(n, c) for n, c in WATCHLIST_STOCKS if c not in ("000660", "005930")]
    STOCKS = _head + _rest
except Exception:
    STOCKS = [("SK하이닉스", "000660"), ("삼성전자", "005930")]
MAX_REPORTS = 6    # 종목별 상세 조회 상한(25종목으로 늘어 속도 고려 — 대형주 외 소형주는 리포트 적음)


def _fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Referer": "https://finance.naver.com/"})
    return urllib.request.urlopen(req, timeout=12).read().decode("euc-kr", errors="replace")


def _price(code):
    """네이버 일봉 최근 종가."""
    try:
        from datetime import date as _d
        end = _d.today().strftime("%Y%m%d")
        url = (f"https://api.finance.naver.com/siseJson.naver?symbol={code}"
               f"&requestType=1&startTime=20260601&endTime={end}&timeframe=day")
        txt = _fetch_raw(url)
        rows = re.findall(r'\["(\d{8})",\s*[\d.]+,\s*[\d.]+,\s*[\d.]+,\s*([\d.]+)', txt)
        return float(rows[-1][1]) if rows else None
    except Exception:
        return None


def _fetch_raw(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Referer": "https://finance.naver.com/"})
    return urllib.request.urlopen(req, timeout=12).read().decode("utf-8", errors="replace")


def _norm_opinion(op):
    o = (op or "").strip()
    if not o or o in ("없음", "N/A", "NR"):
        return "N/A"
    low = o.lower()
    if "buy" in low or "매수" in o or "outperform" in low or "비중확대" in o:
        return "매수"
    if "sell" in low or "매도" in o or "underperform" in low or "비중축소" in o:
        return "매도"
    if "hold" in low or "중립" in o or "보유" in o or "marketperform" in low:
        return "중립"
    return o[:6]


def fetch_reports(code):
    """종목 리서치 목록 + 상세에서 증권사·제목·목표가·투자의견·날짜."""
    out = []
    try:
        lh = _fetch(f"https://finance.naver.com/research/company_list.naver?searchType=itemCode&itemCode={code}")
    except Exception as e:
        print(f"  [WARN] {code} 목록 실패: {e}")
        return out
    for tr in re.findall(r"<tr>(.*?)</tr>", lh, re.S):
        if "company_read" not in tr:
            continue
        nid = re.search(r"nid=(\d+)", tr)
        cells = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S)
        if not nid or len(cells) < 5:
            continue
        title = _html.unescape(re.sub(r"<[^>]+>", "", cells[1])).strip()
        broker = _html.unescape(re.sub(r"<[^>]+>", "", cells[2])).strip()
        date = _html.unescape(re.sub(r"<[^>]+>", "", cells[4])).strip()
        out.append({"nid": nid.group(1), "title": title, "broker": broker, "date": date,
                    "target": None, "opinion": "N/A"})
        if len(out) >= MAX_REPORTS:
            break
    # 상세에서 목표가·투자의견
    for r in out:
        try:
            dh = _fetch(f"https://finance.naver.com/research/company_read.naver?nid={r['nid']}")
            tp = re.search(r'목표가\s*<em[^>]*>\s*<strong>\s*([0-9,]+)', dh)
            op = re.search(r'투자의견\s*<em[^>]*>\s*([^<]+?)\s*</em>', dh)
            if tp:
                r["target"] = int(tp.group(1).replace(",", ""))
            r["opinion"] = _norm_opinion(op.group(1) if op else None)
            time.sleep(0.3)
        except Exception:
            pass
    return out


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    out_stocks = []
    new_alerts = []   # 텔레그램용 신규/목표가변경

    # 직전 상태(중복 방지 + 목표가 변화)
    prev = {}
    try:
        from core import load_state
        prev = load_state("analyst_reports", default={}) or {}
    except Exception:
        pass
    seen = set(prev.get("seen_nids", []))
    prev_targets = prev.get("broker_targets", {})   # f"{code}:{broker}" -> target

    for name, code in STOCKS:
        reports = fetch_reports(code)
        price = _price(code)

        # 파싱 오류 가드: 현재가 대비 비현실적 목표가(0.3~4배 밖)는 오인식으로 보고 제거.
        #   예) 신성이엔지 '목표 2,250'(현재 18,990) — 목표가 아닌 다른 숫자를 잡은 케이스.
        if price:
            for r in reports:
                if r["target"] and not (0.3 * price <= r["target"] <= 4.0 * price):
                    r["target"] = None

        targets = [r["target"] for r in reports if r["target"]]
        consensus = round(sum(targets) / len(targets)) if targets else None
        from collections import Counter
        opc = dict(Counter(r["opinion"] for r in reports if r["opinion"] != "N/A"))

        # 신규 리포트 / 목표가 변경 탐지
        for r in reports:
            if r["nid"] not in seen:
                seen.add(r["nid"])
                tgt = f"{r['target']:,}원" if r["target"] else "—"
                msg = f"{name}: {r['broker']} 목표 {tgt}({r['opinion']}) — {r['title']}"
                # 같은 증권사 직전 목표가 대비 변화
                key = f"{code}:{r['broker']}"
                old = prev_targets.get(key)
                if r["target"] and old and old != r["target"]:
                    arrow = "▲상향" if r["target"] > old else "▼하향"
                    msg += f"  [{arrow} {old:,}→{r['target']:,}]"
                new_alerts.append(msg)
            if r["target"]:
                prev_targets[f"{code}:{r['broker']}"] = r["target"]

        upside = round((consensus / price - 1) * 100, 1) if (consensus and price) else None
        out_stocks.append({
            "name": name, "code": code, "consensus_target": consensus,
            "current_price": round(price) if price else None, "upside_pct": upside,
            "opinion_counts": opc, "n_reports": len(reports), "reports": reports,
        })
        print(f"  {name}: 리포트 {len(reports)}건 · 컨센 목표 {consensus} · 현재 {round(price) if price else '—'} · 상승여력 {upside}% · 의견 {opc}")

    # 텔레그램 경보(신규/목표가변경 — 최대 8건)
    if new_alerts:
        try:
            from core import send_message, get_secret
            if get_secret("TELEGRAM_FINANCE_BOT_TOKEN"):
                body = "📑 국내 증권사 신규 리포트\n" + "\n".join("• " + m for m in new_alerts[:8])
                if len(new_alerts) > 8:
                    body += f"\n…외 {len(new_alerts)-8}건"
                if send_message(body):
                    print(f"  📨 신규 리포트 경보 {len(new_alerts)}건 발송")
        except Exception as e:
            print(f"  [WARN] 텔레그램 실패: {e}")

    # 상태 저장
    try:
        from core import save_state
        save_state("analyst_reports", {"seen_nids": list(seen)[-400:], "broker_targets": prev_targets})
    except Exception:
        pass

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "stocks": out_stocks,
        "note": ("네이버 금융 리서치(무료) — 국내 증권사 종목 리포트: 증권사·제목·목표주가·투자의견·작성일. "
                 "본문(PDF) 전문은 미제공. 목표가 컨센=최근 리포트 목표가 평균. 정보용·투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
