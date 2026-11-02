"""
한국 주식 유니버스 — 시총 상위 N종목 (코스피·코스닥, ETF·ETN 제외)
=====================================================================
네이버 모바일 시총 순위 API 로 종목 코드·이름·시총을 받아 상위 N 개를 돌려준다.
실패 시 직전에 저장한 목록(docs/kr_universe_state.json)을 쓴다.

사용:
    from core.kr_universe import fetch_universe, yf_symbol
    uni = fetch_universe(top_kospi=400, top_kosdaq=400)   # [{code,name,market,mcap_b}]
"""

import json
import os
import re
import sys
import time

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.state_store import load_state, save_state
from market_cap_tracker import NAVER_MV_URL, _diag, _get, _num, is_etf_like

STATE_NAME = "kr_universe"
MIN_UNIVERSE = 300          # 이보다 적게 받으면 네이버 응답이 깨진 것으로 보고 직전 목록 사용


def parse_universe_page(payload):
    """시총 순위 페이지 → [{code, name, mv_eok}] (ETF·ETN 제외, 시총 억원)."""
    items = None
    if isinstance(payload, list):
        items = payload
    elif isinstance(payload, dict):
        for k in ("stocks", "result", "datas", "items", "list"):
            v = payload.get(k)
            if isinstance(v, list):
                items = v
                break
    out = []
    for it in items or []:
        if not isinstance(it, dict) or is_etf_like(it):
            continue
        code = str(it.get("itemCode") or it.get("cd") or it.get("code") or "").strip()
        name = str(it.get("stockName") or it.get("nm") or it.get("name") or "").strip()
        mv = None
        for k in ("marketValue", "marketCap", "mktValue", "marketSum"):
            if it.get(k) is not None:
                mv = _num(it.get(k))
                if mv is not None:
                    break
        if re.fullmatch(r"\d{6}", code) and mv and mv > 0:
            out.append({"code": code, "name": name or code, "mv_eok": mv})
    return out


def _fetch_market(market, top_n):
    rows, seen = [], set()
    for page in range(1, 30):
        text = ""
        try:
            text = _get(NAVER_MV_URL.format(market=market, page=page),
                        headers={"Referer": "https://m.stock.naver.com/", "Accept": "application/json"})
            page_rows = parse_universe_page(json.loads(text))
        except Exception as e:
            print(f"  [WARN] universe {market} p{page}: {type(e).__name__}: {e}")
            if page == 1:
                _diag(f"universe {market}", text)
            break
        new = [r for r in page_rows if r["code"] not in seen]
        if not new:
            break
        seen.update(r["code"] for r in new)
        rows.extend(new)
        if len(rows) >= top_n:
            break
        time.sleep(0.15)
    rows.sort(key=lambda r: -r["mv_eok"])
    return rows[:top_n]


def fetch_universe(top_kospi=400, top_kosdaq=400):
    """[{code, name, market, mcap_b}] — 시총 내림차순. 실패 시 직전 목록."""
    state = load_state(STATE_NAME, {"universe": []})
    uni = []
    for market, n in (("KOSPI", top_kospi), ("KOSDAQ", top_kosdaq)):
        for r in _fetch_market(market, n):
            uni.append({"code": r["code"], "name": r["name"], "market": market, "mcap_b": round(r["mv_eok"] / 10000.0, 3)})
    if len(uni) < MIN_UNIVERSE:
        prev = state.get("universe") or []
        print(f"  [WARN] universe {len(uni)}종목만 수집 → 직전 목록 {len(prev)}종목 사용")
        return prev
    state["universe"] = uni
    state["fetched_at"] = time.strftime("%Y-%m-%d %H:%M")
    save_state(STATE_NAME, state)
    print(f"  유니버스 {len(uni)}종목 (코스피 {sum(1 for u in uni if u['market'] == 'KOSPI')} · 코스닥 {sum(1 for u in uni if u['market'] == 'KOSDAQ')})")
    return uni


def yf_symbol(code, market):
    return f"{code}.{'KS' if market == 'KOSPI' else 'KQ'}"
