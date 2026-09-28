"""
기계 네이티브 경제 모니터 — BlackRock "The Machine-Native Economy"(2026-09) 논지의 라이브 검증
================================================================================================
논지(BlackRock Digital Assets Research): AI=기계 네이티브 지능, 디지털 자산=기계 네이티브 화폐.
에이전트 AI가 스스로 결제·거래하면 (1) 스테이블코인이 거래 화폐로 선도하고 (2) 컴퓨트가
새 디지털 자산 시장이 되며 (3) x402 같은 기계 결제 프로토콜이 확산된다.

이 모듈은 그 세 주장을 매일 측정 가능한 지표와 대조한다. 논문은 '전망'이지 검증된 사실이
아니므로 ai_value_gap.py 와 같은 방식으로 '사실'(라이브 수집)과 '논지'(논문 인용)를 분리해
출력하고, 각 주장마다 검증 상태(측정 / 대리지표 / 소스 미확보)를 명시한다.

측정 지표
  A. 스테이블코인  — CoinGecko 시총·24h 거래량 (1순위, 저장소에서 검증된 경로)
                     → DefiLlama 유통량 (2순위)
     ※ 논문의 '조정 온체인 거래량'(Allium)은 공개 API 가 없다. CoinGecko 24h 거래량은
       거래소 거래량이라 다른 지표다 — 대시보드에 '대리지표'로 표시한다.
  B. 컴퓨트 가격   — vast.ai 마켓플레이스 H100/A100/RTX4090 시간당 임대료 중앙값
                     (논문: "컴퓨트는 새 원자재" → 가격 발견이 시작돼야 한다)
  C. x402 결제     — 공개 통계 엔드포인트 후보를 순회. 확보되면 측정, 아니면 '미확보'로
                     정직하게 표시(논문 스스로 "nascent"라고 인정한 영역).

이력은 docs/machine_economy_state.json 에 하루 1점씩 쌓아(400점 상한) 30/90/365일 변화율과
연환산 성장률을 계산한다 — 논문의 "스테이블코인 80% CAGR" 주장은 이력이 쌓여야 검증된다.
수집 실패 시 직전 값을 유지하고 stale 로 표시, [DIAG] 로 응답 앞부분을 남긴다.

출력: docs/machine_economy.json
🚨 정보 모니터링용. 논문은 BlackRock(IBIT·BUIDL 운용사)의 마케팅 성격 리서치. 투자자문 아님.
"""

import json
import os
import re
import sys
import statistics
import urllib.parse
import urllib.request
from datetime import datetime, timezone, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.state_store import load_state, save_state

KST = timezone(timedelta(hours=9))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "machine_economy.json")
STATE_NAME = "machine_economy"
UA = "Mozilla/5.0 (compatible; ai-finance-machine-economy/1.0)"
MAX_POINTS = 400

STABLE_IDS = ["tether", "usd-coin", "dai", "ethena-usde", "first-digital-usd", "paypal-usd", "usds"]
STABLE_NAMES = {"tether": "USDT", "usd-coin": "USDC", "dai": "DAI", "ethena-usde": "USDe",
                "first-digital-usd": "FDUSD", "paypal-usd": "PYUSD", "usds": "USDS"}
GPUS = [("h100", "H100 SXM"), ("a100", "A100 SXM4"), ("rtx4090", "RTX 4090")]

COINGECKO_MARKETS = ("https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&ids="
                     + ",".join(STABLE_IDS) + "&order=market_cap_desc&per_page=50&page=1")
LLAMA_STABLES = "https://stablecoins.llama.fi/stablecoins?includePrices=false"
VAST_URLS = ["https://console.vast.ai/api/v0/bundles/", "https://cloud.vast.ai/api/v0/bundles/"]
X402_URLS = ["https://www.x402scan.com/api/stats", "https://x402scan.com/api/stats",
             "https://api.x402scan.com/stats", "https://www.x402scan.com/api/v1/stats"]

# ── 논지 (논문 인용 — 라이브 아님) ─────────────────────────────────
THESIS = {
    "source": "BlackRock Digital Assets Research, “The Machine-Native Economy” (2026-09, 11p)",
    "summary": ("AI는 기계 네이티브 지능, 디지털 자산은 기계 네이티브 화폐. 에이전트 AI가 스스로 결제·거래하는 "
                "시대가 오면 둘이 합쳐지며 스테이블코인·블록스페이스·컴퓨트 청구권에 구조적 수요가 생긴다."),
    "pillars": [
        "① LLM 토큰화와 블록체인 토큰화는 '현실 입력을 기계 형식으로 변환'하는 같은 구조 — 에이전트가 온체인 자산을 다루기 쉽다.",
        "② 에이전트 상거래에는 24시간·1센트 미만·초 단위 결제 레일이 필요 — ACH·카드는 부적합, 스테이블코인+x402·MPP·ACP·AP2·TAP 가 대체.",
        "③ 컴퓨트(GPU 시간)가 새 원자재 — 하이퍼스케일러 클라우드 $1.1조(2030), 추론이 최대 워크로드, 컴퓨트 선물·토큰화 청구권 시장 전망.",
    ],
    "paper_figures": [
        {"label": "스테이블코인 시총", "value": "$3,000억+ (2026.9)"},
        {"label": "스테이블코인 조정 거래량", "value": "$11.2조 (2025) — Visa $16.7조·MC $10.6조와 같은 자릿수"},
        {"label": "스테이블코인 거래량 CAGR", "value": "80% (2020→25) vs ACH 8.5%"},
        {"label": "하이퍼스케일러 클라우드 매출", "value": "$1.1조 (2030E, CAGR 29%)"},
        {"label": "AI CapEx 누적", "value": "$5조+ (2025~30)"},
        {"label": "추론 워크로드", "value": "2030년 데이터센터 전력의 43%"},
    ],
    "caveats": [
        "저자 이해관계: BlackRock 은 최대 비트코인 ETF(IBIT)·토큰화 펀드(BUIDL) 운용사 — 자사 상품 수요 논리.",
        "'AI 가 스테이블코인·비트코인을 선호' 근거는 모델 시뮬레이션 응답이지 실제 에이전트 행동이 아님(논문 p.4).",
        "스테이블코인 vs Visa 거래량은 각주에 '직접 비교 불가'(측정 방식 상이).",
        "네이티브 자산(ETH 등) 가치 전이는 '각 네트워크 수수료·스테이킹 설계에 달렸다' — 조건부.",
        "에이전트 결제·컴퓨트 시장은 논문 스스로 'nascent'(초기)라고 인정.",
    ],
}


# ====================================================================
# HTTP · 진단
# ====================================================================
def _get(url, timeout=15):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", errors="replace")


def _diag(label, text):
    body = re.sub(r"\s+", " ", (text or ""))[:300]
    print(f"  [DIAG] {label}: 길이 {len(text or '')} · 앞부분 {body!r}")


def _f(v, default=None):
    try:
        x = float(v)
        return default if x != x else x
    except (TypeError, ValueError):
        return default


# ====================================================================
# 파서 — 순수 함수 (테스트 대상)
# ====================================================================
def parse_coingecko_markets(payload) -> dict:
    """/coins/markets → {"total_cap_b", "total_vol24h_b", "coins":[...]}"""
    coins = []
    if not isinstance(payload, list):
        return {}
    for it in payload:
        if not isinstance(it, dict):
            continue
        cid = it.get("id")
        cap = _f(it.get("market_cap"), 0.0)
        vol = _f(it.get("total_volume"), 0.0)
        if cid and cap > 0:
            coins.append({"id": cid, "name": STABLE_NAMES.get(cid, it.get("symbol", cid).upper()),
                          "cap_b": round(cap / 1e9, 2), "vol24h_b": round(vol / 1e9, 2),
                          "chg_24h_pct": round(_f(it.get("price_change_percentage_24h"), 0.0), 2)})
    if not coins:
        return {}
    coins.sort(key=lambda c: -c["cap_b"])
    return {"total_cap_b": round(sum(c["cap_b"] for c in coins), 2),
            "total_vol24h_b": round(sum(c["vol24h_b"] for c in coins), 2),
            "coins": coins, "source": "coingecko"}


def parse_llama_stables(payload) -> dict:
    """DefiLlama /stablecoins → 총 유통량(USD 페그). 거래량은 없음."""
    assets = (payload or {}).get("peggedAssets") if isinstance(payload, dict) else None
    total = 0.0
    coins = []
    for a in assets or []:
        circ = ((a.get("circulating") or {}).get("peggedUSD")) if isinstance(a, dict) else None
        v = _f(circ, 0.0)
        if v > 0:
            total += v
            coins.append({"id": a.get("gecko_id") or a.get("symbol"), "name": a.get("symbol"),
                          "cap_b": round(v / 1e9, 2), "vol24h_b": None, "chg_24h_pct": None})
    if total <= 0:
        return {}
    coins.sort(key=lambda c: -c["cap_b"])
    return {"total_cap_b": round(total / 1e9, 2), "total_vol24h_b": None,
            "coins": coins[:10], "source": "defillama"}


def parse_vast_offers(payload) -> dict:
    """vast.ai bundles → 시간당 임대료 중앙값·최저·건수."""
    offers = (payload or {}).get("offers") if isinstance(payload, dict) else payload
    prices = []
    for o in offers or []:
        if isinstance(o, dict):
            p = _f(o.get("dph_total") or o.get("dph_base"))
            if p and 0 < p < 500:
                prices.append(p)
    if not prices:
        return {}
    prices.sort()
    return {"median_usd_hr": round(statistics.median(prices), 3), "min_usd_hr": round(prices[0], 3),
            "n_offers": len(prices)}


def parse_x402_stats(payload) -> dict:
    """x402 공개 통계 — 형태를 모르므로 흔한 키 이름을 넓게 받는다."""
    if not isinstance(payload, dict):
        return {}
    flat = {}

    def walk(d, prefix=""):
        for k, v in d.items():
            key = (prefix + "." + k if prefix else k).lower()
            if isinstance(v, dict) and len(flat) < 200:
                walk(v, key)
            else:
                flat[key] = v
    walk(payload)
    out = {}
    for key, v in flat.items():
        base = key.split(".")[-1]
        if base in ("total_transactions", "transactions", "txcount", "tx_count", "totaltransactions", "count") and _f(v) is not None:
            out.setdefault("total_tx", _f(v))
        if base in ("total_volume", "volume", "volume_usd", "totalvolume", "total_volume_usd") and _f(v) is not None:
            out.setdefault("total_volume_usd", _f(v))
        if base in ("transactions_24h", "tx_24h", "daily_transactions", "txs24h") and _f(v) is not None:
            out.setdefault("tx_24h", _f(v))
    return out


# ====================================================================
# 수집기
# ====================================================================
def fetch_stablecoins() -> dict:
    text = ""
    try:
        text = _get(COINGECKO_MARKETS)
        r = parse_coingecko_markets(json.loads(text))
        if r:
            return r
        _diag("coingecko markets (0건)", text)
    except Exception as e:
        print(f"  [WARN] coingecko: {type(e).__name__}: {e}")
        if text:
            _diag("coingecko", text)
    text = ""
    try:
        text = _get(LLAMA_STABLES, timeout=20)
        r = parse_llama_stables(json.loads(text))
        if r:
            return r
        _diag("defillama (0건)", text)
    except Exception as e:
        print(f"  [WARN] defillama: {type(e).__name__}: {e}")
        if text:
            _diag("defillama", text)
    return {}


def fetch_compute() -> dict:
    out = {}
    for key, gpu in GPUS:
        q = json.dumps({"gpu_name": {"eq": gpu}, "rentable": {"eq": True}, "num_gpus": {"eq": 1},
                        "order": [["dph_total", "asc"]], "limit": 200}, separators=(",", ":"))
        got = None
        for base in VAST_URLS:
            text = ""
            try:
                text = _get(base + "?q=" + urllib.parse.quote(q))
                r = parse_vast_offers(json.loads(text))
                if r:
                    got = r
                    break
                _diag(f"vast {gpu} (0건)", text)
            except Exception as e:
                print(f"  [WARN] vast {gpu} @ {base.split('/')[2]}: {type(e).__name__}: {e}")
                if text:
                    _diag(f"vast {gpu}", text)
        if got:
            got["gpu"] = gpu
            out[key] = got
    if out:
        out["source"] = "vast.ai"
    return out


def fetch_x402() -> dict:
    for url in X402_URLS:
        text = ""
        try:
            text = _get(url, timeout=12)
            r = parse_x402_stats(json.loads(text))
            if r:
                r["source"] = url.split("/")[2]
                return r
            _diag(f"x402 {url.split('/')[2]} (키 미매칭)", text)
        except Exception as e:
            print(f"  [WARN] x402 {url}: {type(e).__name__}: {e}")
            if text:
                _diag("x402", text)
    return {}


# ====================================================================
# 이력 · 파생 · 판정 (순수 함수)
# ====================================================================
def merge_history(series: list, point: dict) -> list:
    """같은 날짜는 교체, 날짜순 정렬, 상한 유지."""
    s = [p for p in (series or []) if isinstance(p, dict) and p.get("date") != point.get("date")]
    s.append(point)
    s.sort(key=lambda p: p.get("date", ""))
    return s[-MAX_POINTS:]


def _at_or_before(series, date_str):
    cand = [p for p in series if p.get("date", "") <= date_str]
    return cand[-1] if cand else None


def derive_trend(series: list, key: str, today: str) -> dict:
    """30/90/365일 변화율 + 연환산 성장률(최소 60일 이력)."""
    cur = _at_or_before(series, today)
    if not cur or _f(cur.get(key)) is None:
        return {}
    v0 = _f(cur[key])
    out = {"current": v0, "days_of_history": 0}
    if v0 is None or v0 <= 0:
        return out
    d = datetime.strptime(today, "%Y-%m-%d")
    first = next((p for p in series if _f(p.get(key)) is not None and _f(p.get(key)) > 0), None)
    if first:
        out["days_of_history"] = (d - datetime.strptime(first["date"], "%Y-%m-%d")).days
    for label, days in (("chg_30d_pct", 30), ("chg_90d_pct", 90), ("chg_365d_pct", 365)):
        ref = _at_or_before(series, (d - timedelta(days=days)).strftime("%Y-%m-%d"))
        rv = _f(ref.get(key)) if ref else None
        # 기준점이 실제로 그만큼 오래됐을 때만 (이력 초기에 30일 전 값 = 어제 값 이 되는 오류 방지)
        if ref and rv and rv > 0 and (d - datetime.strptime(ref["date"], "%Y-%m-%d")).days >= days * 0.8:
            out[label] = round((v0 / rv - 1) * 100, 2)
    if first and out["days_of_history"] >= 60:
        fv = _f(first[key])
        yrs = out["days_of_history"] / 365.25
        if fv and fv > 0 and yrs > 0:
            out["annualized_growth_pct"] = round(((v0 / fv) ** (1 / yrs) - 1) * 100, 1)
    return out


def build_claims(stable: dict, compute: dict, x402: dict, trend: dict) -> list:
    """논문 주장별 검증 상태·판정. status: measured | proxy | unavailable"""
    claims = []

    # ① 스테이블코인 성장 (논문: 거래량 CAGR 80%)
    st = trend.get("stable_cap_b") or {}
    if stable.get("total_cap_b"):
        g = st.get("annualized_growth_pct")
        if g is not None:
            verdict = ("논문 속도(80%)에 근접" if g >= 60 else
                       "성장 중이나 논문 속도 미달" if g >= 20 else
                       "정체" if g > -5 else "축소")
            detail = f"시총 연환산 {g:+.1f}% ({st['days_of_history']}일 이력)"
        else:
            verdict = "이력 축적 중 (60일 후 연환산 산출)"
            detail = f"시총 ${stable['total_cap_b']:,.0f}B" + (f" · 30일 {st['chg_30d_pct']:+.1f}%" if st.get("chg_30d_pct") is not None else "")
        claims.append({"claim": "스테이블코인이 기계 화폐로 빠르게 확산 (거래량 CAGR 80%)",
                       "paper": "$3,000억+ 시총 · $11.2조 조정거래량(2025) · CAGR 80%",
                       "status": "proxy", "verdict": verdict, "detail": detail,
                       "note": "논문 지표(Allium 조정 온체인 거래량)는 공개 API 없음 → 시총 성장률·거래소 24h 거래량으로 대리 측정"})
    else:
        claims.append({"claim": "스테이블코인이 기계 화폐로 빠르게 확산", "paper": "CAGR 80%",
                       "status": "unavailable", "verdict": "수집 실패", "detail": "CoinGecko·DefiLlama 모두 실패", "note": ""})

    # ② 컴퓨트 = 새 원자재 (가격 발견)
    ct = trend.get("h100_usd_hr") or {}
    if compute.get("h100"):
        h = compute["h100"]
        detail = f"H100 ${h['median_usd_hr']:.2f}/h"
        if h.get("min_usd_hr") is not None and h.get("n_offers"):
            detail += f" (최저 ${h['min_usd_hr']:.2f}, {h['n_offers']}건)"
        elif str(compute.get("source", "")).startswith("stale"):
            detail += " (직전 값 유지)"
        if ct.get("chg_90d_pct") is not None:
            detail += f" · 90일 {ct['chg_90d_pct']:+.1f}%"
            verdict = "가격 하락 — 공급 확대/세대 교체" if ct["chg_90d_pct"] < -10 else \
                      "가격 상승 — 수요 초과" if ct["chg_90d_pct"] > 10 else "가격 안정"
        else:
            verdict = "가격 발견 진행 중 (마켓 시세 확보)"
        claims.append({"claim": "컴퓨트(GPU 시간)가 가격 발견되는 새 원자재가 된다",
                       "paper": "클라우드 $1.1조(2030) · 컴퓨트 선물·토큰화 청구권 전망",
                       "status": "measured", "verdict": verdict, "detail": detail,
                       "note": "vast.ai 마켓플레이스 현물 임대료 — 선물·토큰화 시장은 아직 없음(논문도 '설계 과제' 인정)"})
    else:
        claims.append({"claim": "컴퓨트가 새 원자재가 된다", "paper": "클라우드 $1.1조(2030)",
                       "status": "unavailable", "verdict": "수집 실패", "detail": "vast.ai 응답 없음 — [DIAG] 확인", "note": ""})

    # ③ x402 기계 결제 확산
    if x402.get("total_tx") or x402.get("total_volume_usd"):
        detail = " · ".join(filter(None, [
            f"누적 {x402['total_tx']:,.0f}건" if x402.get("total_tx") else None,
            f"${x402['total_volume_usd']/1e6:,.1f}M" if x402.get("total_volume_usd") else None,
            f"24h {x402['tx_24h']:,.0f}건" if x402.get("tx_24h") else None]))
        claims.append({"claim": "x402 등 기계 결제 프로토콜이 확산된다", "paper": "x402(Coinbase)·MPP·ACP·AP2·TAP",
                       "status": "measured", "verdict": "측정 시작", "detail": detail, "note": x402.get("source", "")})
    else:
        claims.append({"claim": "x402 등 기계 결제 프로토콜이 확산된다", "paper": "x402(Coinbase)·MPP·ACP·AP2·TAP",
                       "status": "unavailable", "verdict": "검증 불가 — 공개 통계 소스 미확보",
                       "detail": "논문도 'nascent' 인정. Dune/CDP API 키 등록 시 측정 가능",
                       "note": "후보 엔드포인트 응답 없음 — [DIAG] 참조"})
    return claims


# ====================================================================
# 메인
# ====================================================================
def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    now = datetime.now(KST)
    today = now.strftime("%Y-%m-%d")
    print("=" * 60)
    print("  기계 네이티브 경제 모니터 (BlackRock 논지 검증)")
    print(f"  KST: {now:%Y-%m-%d %H:%M:%S}")
    print("=" * 60)

    state = load_state(STATE_NAME, {"series": []})
    series = state.get("series") or []
    last = series[-1] if series else {}

    stable = fetch_stablecoins()
    compute = fetch_compute()
    x402 = fetch_x402()
    stale = []

    # 실패 시 직전 값 유지 (단위 섞임·빈 카드보다 낫다) — stale 로 표시
    if not stable and last.get("stable_cap_b"):
        stable = {"total_cap_b": last["stable_cap_b"], "total_vol24h_b": last.get("stable_vol_b"),
                  "coins": [], "source": f"stale({last.get('date')})"}
        stale.append("stablecoins")
    if not compute and last.get("h100_usd_hr"):
        compute = {"h100": {"gpu": "H100 SXM", "median_usd_hr": last["h100_usd_hr"], "min_usd_hr": None, "n_offers": 0},
                   "source": f"stale({last.get('date')})"}
        stale.append("compute")

    print(f"  스테이블코인: {'$%.0fB' % stable['total_cap_b'] if stable.get('total_cap_b') else '실패'} ({stable.get('source', '-')})")
    print(f"  컴퓨트: " + (", ".join(f"{k.upper()} ${compute[k]['median_usd_hr']:.2f}/h" for k, _ in GPUS if k in compute) or "실패"))
    print(f"  x402: {x402 or '미확보'}")

    point = {"date": today,
             "stable_cap_b": stable.get("total_cap_b"), "stable_vol_b": stable.get("total_vol24h_b"),
             "h100_usd_hr": (compute.get("h100") or {}).get("median_usd_hr"),
             "a100_usd_hr": (compute.get("a100") or {}).get("median_usd_hr"),
             "rtx4090_usd_hr": (compute.get("rtx4090") or {}).get("median_usd_hr"),
             "x402_tx": x402.get("total_tx")}
    # 새로 수집된 값이 하나라도 있을 때만 이력에 쌓는다 — 전부 실패/stale 이면 빈 점을 남기지 않는다
    has_fresh = ((bool(stable) and "stablecoins" not in stale) or
                 (bool(compute) and "compute" not in stale) or bool(x402))
    if has_fresh:
        series = merge_history(series, point)
    state["series"] = series
    state["last_run"] = now.isoformat()
    save_state(STATE_NAME, state)

    trend = {k: derive_trend(series, k, today) for k in ("stable_cap_b", "stable_vol_b", "h100_usd_hr", "a100_usd_hr", "rtx4090_usd_hr")}
    claims = build_claims(stable, compute, x402, trend)

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "facts": {"stablecoins": stable, "compute": compute, "x402": x402, "stale": stale},
        "trend": trend,
        "history_points": len(series),
        "history_tail": series[-90:],
        "claims": claims,
        "thesis": THESIS,
        "note": ("'사실'은 매일 라이브 수집(CoinGecko·vast.ai), '논지'는 BlackRock 논문 인용. "
                 "측정=논문 지표 직접 측정 · 대리=다른 지표로 근사 · 미확보=공개 소스 없음. 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"[OK] {OUTPUT_FILE}  이력 {len(series)}점 · stale {stale or '없음'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
