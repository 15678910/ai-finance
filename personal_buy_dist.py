"""
개인 매수 가격대 분포 — 관심종목 개인 순매수의 가격선 분포도 (매물대 추정)
==========================================================================
투자자별 매매동향(개인 순매수량, KRX)과 일별 종가를 결합해
'개인 투자자가 최근 1년간 어느 가격대에서 얼마나 순매수했는지' 분포를 만든다.
  · 가격 버킷별: 매수우위일 수량 합(buy) / 매도우위일 수량 합(sell) / 순합(net)
  · 개인 추정 평균 매수단가 = Σ(종가×순매수량, 순매수일만) / Σ(순매수량)
  · 현재가 위쪽 버킷의 순매수 물량 비율 = '물린 물량' 추정 (잠재 매물대/저항)
※ KRX가 데이터 API에 로그인 요구 + GitHub IP 차단 → 내 PC(한국 IP) 로컬 스케줄러 전용.
※ 근사치: 일별 순매수×종가 기반 추정이며 실제 개인 보유단가와 다를 수 있음.

출력: docs/personal_buy_dist.json
🚨 정보 제공용 · 투자자문 아님.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "personal_buy_dist.json")

WATCH = [
    ("KOSPI", "KOSPI"),                  # 시장 전체 개인 순매수 × 지수 종가 (특수 처리)
    ("000660", "SK하이닉스"),
    ("108490", "로보티즈"),
    ("005930", "삼성전자"),
    ("003550", "LG"),
    ("066570", "LG전자"),
    ("042700", "한미반도체"),
]
N_BUCKETS = 16
LOOKBACK_DAYS = 370                      # 최근 1년(영업일 ~245일)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    if not (os.environ.get("KRX_ID") and os.environ.get("KRX_PW")):
        print("[INFO] KRX_ID/KRX_PW 미설정 — 개인 매수 분포는 KRX 로그인 필수라 수집 불가. 기존 파일 보존.")
        return 0

    try:
        from pykrx import stock
    except Exception as e:
        print(f"[ERROR] pykrx import 실패: {e}")
        return 1

    # 로그인 사전 점검 (반복 실패로 인한 계정 잠금 방지)
    try:
        from pykrx.website.comm.auth import login_krx
        if not login_krx(os.environ["KRX_ID"], os.environ["KRX_PW"]):
            print("[ERROR] KRX 로그인 거부 — 중단(기존 파일 보존).")
            return 1
    except ImportError:
        pass

    now = datetime.now(KST)
    frm = (now - timedelta(days=LOOKBACK_DAYS)).strftime("%Y%m%d")
    to = now.strftime("%Y%m%d")

    stocks, asof = [], None
    for code, name in WATCH:
        try:
            if code == "KOSPI":                                            # 시장 전체 + 지수 종가
                vol = stock.get_market_trading_volume_by_date(frm, to, "KOSPI")
                ohl = stock.get_index_ohlcv_by_date(frm, to, "1001")
            else:
                vol = stock.get_market_trading_volume_by_date(frm, to, code)   # 투자자별 순매수량
                ohl = stock.get_market_ohlcv_by_date(frm, to, code)            # 일별 종가
        except Exception as e:
            print(f"  [WARN] {name}({code}) 조회 실패: {e}")
            continue
        if vol is None or len(vol) == 0 or ohl is None or len(ohl) == 0 or "개인" not in vol.columns:
            print(f"  [WARN] {name}({code}) 데이터 없음")
            continue

        joined = []
        closes = ohl["종가"].to_dict()
        for idx, row in vol.iterrows():
            c = closes.get(idx)
            if c is None or c <= 0:
                continue
            joined.append((float(c), int(row["개인"])))
        if len(joined) < 60:
            print(f"  [WARN] {name}({code}) 표본 부족({len(joined)}일)")
            continue

        last_dt = ohl.index[-1]
        asof = last_dt.strftime("%Y-%m-%d") if hasattr(last_dt, "strftime") else str(last_dt)[:10]
        cur = float(ohl["종가"].iloc[-1])
        lo = min(p for p, _ in joined)
        hi = max(p for p, _ in joined)
        span = (hi - lo) or 1.0

        buckets = [{"lo": lo + span * i / N_BUCKETS, "hi": lo + span * (i + 1) / N_BUCKETS,
                    "buy": 0, "sell": 0} for i in range(N_BUCKETS)]
        buy_val = buy_vol = 0
        for p, v in joined:
            bi = min(N_BUCKETS - 1, int((p - lo) / span * N_BUCKETS))
            if v >= 0:
                buckets[bi]["buy"] += v
                buy_val += p * v
                buy_vol += v
            else:
                buckets[bi]["sell"] += -v
        for b in buckets:
            b["net"] = b["buy"] - b["sell"]
            b["lo"] = round(b["lo"])
            b["hi"] = round(b["hi"])

        avg_buy = round(buy_val / buy_vol) if buy_vol > 0 else None
        net_total = sum(b["net"] for b in buckets)
        # 현재가 위쪽 버킷의 양(+)의 net 물량 비율 = '물린 물량' 추정
        above_net = sum(max(0, b["net"]) for b in buckets if b["lo"] >= cur)
        pos_net = sum(max(0, b["net"]) for b in buckets)
        above_frac = round(above_net / pos_net * 100) if pos_net > 0 else None

        stocks.append({
            "code": code, "name": name, "asof": asof, "price": round(cur),
            "lookback_days": len(joined),
            "buckets": buckets, "avg_buy_price": avg_buy,
            "net_total": net_total,
            "above_frac": above_frac,
            "vs_avg_pct": round((cur / avg_buy - 1) * 100, 1) if avg_buy else None,
        })
        print(f"  {name}: 현재 {cur:,.0f} · 개인 추정평단 {avg_buy:,} ({round((cur/avg_buy-1)*100,1) if avg_buy else '—'}%) "
              f"· 1년 순매수 {net_total:+,}주 · 물린물량 {above_frac}%")

    if not stocks:
        print("[ERROR] 전 종목 실패 — 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": asof,
        "stocks": stocks,
        "note": ("개인 투자자 순매수량(KRX 투자자별 매매동향)×일별 종가로 최근 1년 가격대별 분포를 추정. "
                 "buy=매수우위일 수량합, sell=매도우위일 수량합, net=순합. "
                 "평균 매수단가=순매수일 가중평균(근사). 현재가 위 net+물량=개인 '물린' 추정 구간(잠재 매물대). "
                 "실제 개인 보유단가와 다를 수 있는 근사치 · 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE} ({len(stocks)}종목 · 기준 {asof})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
