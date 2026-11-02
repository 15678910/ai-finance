"""
수급 연속성 스크리너 — 외국인·기관 연속 순매수 종목 (시총 상위 유니버스)
=========================================================================
브리핑의 '수급 연속 노드'를 대시보드용으로 옮긴 것:
  • 외국인 또는 기관이 3일 이상 연속 순매수한 종목만 남긴다
  • 동반 순매수(외국인·기관 모두) 연속일이 길수록 앞에 둔다 → 같으면 더 긴 연속일 → 같으면 5일 순매수 금액
  • 별도 목록: 동반 순매수 2일 이상(데이트레이더 노드), 동반 연속 순매도 3일 이상(경고)
  • 시장폭 게이트(docs/market_breadth.json)가 '차단'이면 결과에 gate_blocked 표시 — 화면에서 회색 처리

데이터: 네이버 종목별 외국인·기관 매매 페이지(investor_flow_monitor.fetch_naver_flow 재사용, 최근 20거래일).
유니버스: 코스피·코스닥 시총 상위 각 400 (core.kr_universe). 수집률 50% 미만이면 덮어쓰지 않고 실패(exit 1).
출력: docs/flow_streak.json
"""

import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
from core.kr_universe import fetch_universe
from investor_flow_monitor import fetch_naver_flow as fetch_naver_flow_html
from market_cap_tracker import _diag, _get

KST = timezone(timedelta(hours=9))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "flow_streak.json")
BREADTH_FILE = os.path.join(BASE_DIR, "docs", "market_breadth.json")
TOP_KOSPI, TOP_KOSDAQ = 400, 400
MIN_STREAK = 3           # 외국인 또는 기관 연속 순매수 최소 일수
JOINT_MIN = 2            # 동반 순매수 목록 최소 일수
WORKERS = 8
MIN_COVERAGE = 0.5
TOP_N = 80
# 첫 실행(2026-11-02): finance.naver.com PC 페이지가 Actions 에서 0건 → 모바일 JSON API 를 먼저 쓰고 PC 페이지는 최후 수단
TREND_URLS = ("https://m.stock.naver.com/api/stock/{code}/trend?pageSize=20&page=1",
              "https://api.stock.naver.com/stock/{code}/trend?pageSize=20&page=1")
_DIAG_ONCE = {"done": False}


def _num(v):
    try:
        return float(str(v).replace(",", "").replace("+", "").strip())
    except (TypeError, ValueError):
        return 0.0


def parse_trend_json(payload):
    """모바일 투자자 동향 API → fetch_naver_flow 와 같은 행 형식 (날짜 내림차순).
    필드명 변종을 모두 받는다: bizdate/date, closePrice, foreignerPureBuyQuant/foreignNetBuy..., organPureBuyQuant/institution..."""
    items = payload
    if isinstance(payload, dict):
        for k in ("result", "datas", "items", "list", "trend", "trends"):
            if isinstance(payload.get(k), list):
                items = payload[k]
                break
    if not isinstance(items, list):
        return []
    rows = []
    for it in items:
        if not isinstance(it, dict):
            continue
        d = str(it.get("bizdate") or it.get("date") or it.get("localTradedAt") or it.get("tradeDate") or "")
        d = re.sub(r"[^0-9]", "", d)
        if len(d) < 8:
            continue
        d = f"{d[:4]}-{d[4:6]}-{d[6:8]}"
        fk = next((k for k in it if re.match(r"(foreigner|foreign|frgn)[A-Za-z]*(PureBuy|NetBuy|netBuy|Net)", k)), None)
        ik = next((k for k in it if re.match(r"(organ|institution|inst)[A-Za-z]*(PureBuy|NetBuy|netBuy|Net)", k)), None)
        if not fk or not ik:
            continue
        chg = it.get("fluctuationsRatio") or it.get("changeRate") or it.get("change_pct") or 0
        rows.append({"date": d, "close": int(_num(it.get("closePrice") or it.get("close") or 0)),
                     "change_pct": _num(chg), "volume": int(_num(it.get("accumulatedTradingVolume") or it.get("volume") or 0)),
                     "institutional_net_shares": int(_num(it.get(ik))), "foreign_net_shares": int(_num(it.get(fk))),
                     "foreign_holding_pct": _num(it.get("foreignerHoldRatio") or it.get("foreignHoldRatio") or 0)})
    rows.sort(key=lambda r: r["date"], reverse=True)
    # 등락률 필드명을 못 맞춘 경우(첫 실행: 전부 0.0%) 전일 종가로 직접 계산
    for a, b in zip(rows, rows[1:]):
        if not a["change_pct"] and a["close"] and b["close"]:
            a["change_pct"] = round((a["close"] / b["close"] - 1) * 100, 2)
    return rows


def fetch_naver_flow(code):
    """모바일 JSON API(2종) → 실패 시 PC HTML 페이지. 첫 실패 응답은 [DIAG] 로 한 번 남긴다."""
    for url in TREND_URLS:
        text = ""
        try:
            text = _get(url.format(code=code), headers={"Referer": "https://m.stock.naver.com/", "Accept": "application/json"})
            rows = parse_trend_json(json.loads(text))
            if rows:
                return rows
        except Exception as e:
            rows = []
            err = f"{type(e).__name__}: {e}"
        else:
            err = "0행"
        if not _DIAG_ONCE["done"]:
            _DIAG_ONCE["done"] = True
            print(f"  [WARN] trend {code} @ {url.split('/')[2]}: {err}")
            _diag(f"trend {code}", text)
    return fetch_naver_flow_html(code)


# ====================================================================
# 순수 계산 (테스트 대상)
# ====================================================================
def streak(values):
    """최신순 값 리스트 → 연속 양수면 +n, 연속 음수면 -n, 첫 값이 0이면 0."""
    if not values:
        return 0
    sign = 1 if values[0] > 0 else -1 if values[0] < 0 else 0
    if sign == 0:
        return 0
    n = 0
    for v in values:
        if (v > 0) == (sign > 0) and v != 0:
            n += 1
        else:
            break
    return n * sign


def compute_streaks(rows):
    """fetch_naver_flow 행(날짜 내림차순 보장) → 연속일·5일 금액(억원)·당일 등락."""
    rows = sorted([r for r in rows if r.get("date")], key=lambda r: r["date"], reverse=True)
    if not rows:
        return {}
    f = [r.get("foreign_net_shares", 0) for r in rows]
    i = [r.get("institutional_net_shares", 0) for r in rows]
    joint = 0
    for a, b in zip(f, i):
        if a > 0 and b > 0:
            joint += 1
        else:
            break
    joint_sell = 0
    for a, b in zip(f, i):
        if a < 0 and b < 0:
            joint_sell += 1
        else:
            break
    eok = lambda vals: round(sum(v * r.get("close", 0) for v, r in zip(vals, rows[:5])) / 1e8, 1)
    return {"as_of": rows[0]["date"], "close": rows[0].get("close"), "chg_pct": rows[0].get("change_pct"),
            "f_streak": streak(f), "i_streak": streak(i), "joint": joint, "joint_sell": joint_sell,
            "f5_eok": eok(f), "i5_eok": eok(i), "days": len(rows)}


def screen(universe, flows):
    """universe [{code,name,market,mcap_b}] + flows {code: rows} → 매수·동반·매도 목록."""
    items = []
    for u in universe:
        rows = flows.get(u["code"])
        if not rows:
            continue
        s = compute_streaks(rows)
        if not s:
            continue
        items.append({**u, **s})
    buy = [x for x in items if x["f_streak"] >= MIN_STREAK or x["i_streak"] >= MIN_STREAK]
    buy.sort(key=lambda x: (-x["joint"], -max(x["f_streak"], x["i_streak"]), -(x["f5_eok"] + x["i5_eok"])))
    joint2 = [x for x in items if x["joint"] >= JOINT_MIN]
    joint2.sort(key=lambda x: (-x["joint"], -(x["f5_eok"] + x["i5_eok"])))
    sell = [x for x in items if x["joint_sell"] >= MIN_STREAK]
    sell.sort(key=lambda x: (-x["joint_sell"], (x["f5_eok"] + x["i5_eok"])))
    as_of = max((x["as_of"] for x in items), default=None)
    return {"as_of": as_of,
            "counts": {"universe": len(universe), "fetched": len(items), "buy": len(buy), "joint2": len(joint2), "sell": len(sell)},
            "buy": buy[:TOP_N], "joint2": joint2[:40], "sell": sell[:20]}


def load_gate():
    try:
        with open(BREADTH_FILE, encoding="utf-8") as f:
            b = json.load(f)
        g = b.get("gate") or {}
        return {"open": g.get("open"), "score": b.get("score"), "grade": b.get("grade"), "reason": g.get("reason"), "as_of": b.get("as_of")}
    except Exception:
        return None


# ====================================================================
# 수집
# ====================================================================
def fetch_all_flows(universe):
    def one(u):
        for attempt in range(2):
            rows = fetch_naver_flow(u["code"])
            if rows:
                return u["code"], rows
            time.sleep(0.5)
        return u["code"], []
    flows = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        for code, rows in ex.map(one, universe):
            if rows:
                flows[code] = rows
    return flows


def main():
    now = datetime.now(KST)
    print("=" * 60)
    print("  수급 연속성 스크리너 (외국인·기관 연속 순매수)")
    print(f"  KST: {now:%Y-%m-%d %H:%M:%S}")
    print("=" * 60)
    universe = fetch_universe(TOP_KOSPI, TOP_KOSDAQ)
    if not universe:
        print("  [FAIL] 유니버스 없음")
        return 1
    t0 = time.time()
    flows = fetch_all_flows(universe)
    cov = len(flows) / len(universe)
    print(f"  수급 수집: {len(flows)}/{len(universe)} ({cov:.0%}, {time.time() - t0:.0f}s)")
    if cov < MIN_COVERAGE:
        print(f"  [FAIL] 수집률 {cov:.0%} < {MIN_COVERAGE:.0%} → 기존 결과 유지, 실패 처리")
        return 1
    res = screen(universe, flows)
    gate = load_gate()
    blocked = bool(gate and gate.get("open") is False)
    c = res["counts"]
    print(f"  기준일 {res['as_of']} · 연속 순매수 {c['buy']} · 동반 2일+ {c['joint2']} · 동반 순매도 {c['sell']} · 게이트 {'차단' if blocked else '열림' if gate else '정보 없음'}")
    for x in res["buy"][:5]:
        print(f"    {x['name']:10} 외인 {x['f_streak']:+d}일 기관 {x['i_streak']:+d}일 동반 {x['joint']}일 · 5일 {x['f5_eok'] + x['i5_eok']:+.0f}억")
    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "as_of": res["as_of"],
        "rule": f"외국인 또는 기관 {MIN_STREAK}일+ 연속 순매수 → 동반 연속일 → 최장 연속일 → 5일 순매수 금액 순",
        "gate": gate, "gate_blocked": blocked,
        "counts": c, "buy": res["buy"], "joint2": res["joint2"], "sell": res["sell"],
        "note": ("네이버 종목별 매매 페이지 최근 20거래일 기준(연속일 최대 20). 금액은 주식수×종가 근사(억원). "
                 "시총 상위 800종목만 본다. 수급은 보조 신호이며 매수 추천이 아니다."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
