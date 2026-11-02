"""
네이버 모바일 '투자자 동향' API — 종목별 일별 외국인·기관 순매수 (공용)
======================================================================
finance.naver.com PC 페이지(frgn.naver)는 2026-10-30 이후 GitHub Actions 에서 데이터 행을
돌려주지 않는다. 모바일 JSON API 는 2026-11-02 수급 스크리너에서 800/800 종목 수집을 확인했다.

    from core.naver_flow import fetch_trend_rows
    rows = fetch_trend_rows("005930")   # 날짜 내림차순, investor_flow_monitor 와 같은 행 형식

행 형식: {date, close, change_pct, volume, institutional_net_shares, foreign_net_shares,
         foreign_holding_shares, foreign_holding_pct}
"""

import json
import re
import sys
import os

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
from market_cap_tracker import _diag, _get   # noqa: E402

TREND_URLS = ("https://m.stock.naver.com/api/stock/{code}/trend?pageSize={n}&page=1",
              "https://api.stock.naver.com/stock/{code}/trend?pageSize={n}&page=1")
DIAG_ONCE = {"fail": False, "keys": False}


def _num(v):
    try:
        return float(str(v).replace(",", "").replace("+", "").strip())
    except (TypeError, ValueError):
        return 0.0


def parse_trend_json(payload):
    """모바일 투자자 동향 API → 행 리스트(날짜 내림차순). 필드명 변종을 모두 받는다:
    bizdate/date, closePrice, foreignerPureBuyQuant/foreignNetBuy..., organPureBuyQuant/institution...,
    foreignerHoldRatio/foreignHoldRatio. 등락률 필드가 없으면 전일 종가로 계산."""
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
        hk = next((k for k in it if re.match(r"(foreigner|foreign|frgn)[A-Za-z]*(HoldRatio|HoldingRatio|Ratio|holdRate)", k)), None)
        hs = next((k for k in it if re.match(r"(foreigner|foreign|frgn)[A-Za-z]*(HoldQuant|HoldingQuant|HoldShares|HoldQty)", k)), None)
        chg = it.get("fluctuationsRatio") or it.get("changeRate") or it.get("change_pct") or 0
        rows.append({"date": d, "close": int(_num(it.get("closePrice") or it.get("close") or 0)),
                     "change_pct": _num(chg), "volume": int(_num(it.get("accumulatedTradingVolume") or it.get("volume") or 0)),
                     "institutional_net_shares": int(_num(it.get(ik))), "foreign_net_shares": int(_num(it.get(fk))),
                     "foreign_holding_shares": int(_num(it.get(hs))) if hs else 0,
                     "foreign_holding_pct": _num(it.get(hk)) if hk else 0.0})
        if not DIAG_ONCE["keys"]:
            DIAG_ONCE["keys"] = True
            print(f"  [DIAG] trend 응답 키: {sorted(it.keys())[:30]}")
    rows.sort(key=lambda r: r["date"], reverse=True)
    for a, b in zip(rows, rows[1:]):
        if not a["change_pct"] and a["close"] and b["close"]:
            a["change_pct"] = round((a["close"] / b["close"] - 1) * 100, 2)
    return rows


def fetch_trend_rows(code, n=20, get=None):
    """모바일 JSON API 2개 호스트 순서로 시도. 전부 실패하면 [] (첫 실패 응답은 [DIAG] 한 번)."""
    get = get or _get
    for url in TREND_URLS:
        text, err = "", ""
        try:
            text = get(url.format(code=code, n=n), headers={"Referer": "https://m.stock.naver.com/", "Accept": "application/json"})
            rows = parse_trend_json(json.loads(text))
            if rows:
                return rows
            err = "0행"
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
        if not DIAG_ONCE["fail"]:
            DIAG_ONCE["fail"] = True
            print(f"  [WARN] trend {code} @ {url.split('/')[2]}: {err}")
            _diag(f"trend {code}", text)
    return []
