"""
아시아 반도체 섹터 당일 모니터 — KOSPI 반도체 '디커플링'의 선행 동력 실시간 관측
=================================================================================
한국 반도체(삼성·SK) 장중 급등의 진짜 동력은 '한국만의 수급'이 아니라
같은 시간대 거래되는 아시아 반도체(TSMC·도쿄일렉트론·어드반테스트)의 동반 흐름인 경우가 많다.
개장 전 미국 선물로는 못 잡지만, KOSPI 장중(09:00~15:30 KST)과 겹치는
대만·일본 반도체는 '당일 실시간' 관측이 가능하다.
→ '아시아 동조 랠리' vs '한국 단독'을 실시간 판별해 예측 미스의 원인을 드러낸다.

출력: docs/asia_semi.json
🚨 통계·시세 관측. 투자 결정 단독 사용 금지.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "asia_semi.json")

# (표시명, 티커, 국기, 그룹) — 그룹: korea / asia / us
UNIVERSE = [
    ("SK하이닉스", "000660.KS", "🇰🇷", "korea"),
    ("삼성전자", "005930.KS", "🇰🇷", "korea"),
    ("TSMC", "2330.TW", "🇹🇼", "asia"),
    ("도쿄일렉트론", "8035.T", "🇯🇵", "asia"),
    ("어드반테스트", "6857.T", "🇯🇵", "asia"),
    ("마이크론(메모리)", "MU", "🇺🇸", "us"),
    ("브로드컴(AI)", "AVGO", "🇺🇸", "us"),
    ("SOXX(섹터)", "SOXX", "🇺🇸", "us"),
]


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    import yfinance as yf
    import warnings
    warnings.filterwarnings("ignore")

    now = datetime.now(KST)
    print("=" * 55)
    print("  아시아 반도체 섹터 당일 모니터")
    print(f"  KST: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    items = []
    for name, tk, flag, grp in UNIVERSE:
        try:
            h = yf.Ticker(tk).history(period="6d")
            c = h["Close"].dropna()
            if len(c) < 2:
                print(f"  [THIN] {name} ({tk})")
                continue
            last, prev = float(c.iloc[-1]), float(c.iloc[-2])
            chg = (last / prev - 1) * 100
            # 당일 장중 진폭(고가-저가 / 시가) — 변동성 참고
            row = h.iloc[-1]
            intraday = None
            try:
                intraday = round((float(row["High"]) / float(row["Open"]) - 1) * 100, 2)
            except Exception:
                pass
            items.append({
                "name": name, "ticker": tk, "flag": flag, "group": grp,
                "last": round(last, 2), "chg_pct": round(chg, 2),
                "intraday_up_pct": intraday,
                "asof": str(c.index[-1].date()),
            })
            print(f"  {flag} {name:16s} {chg:+.2f}%  (asof {c.index[-1].date()})")
        except Exception as e:
            print(f"  [ERR] {name} ({tk}): {e}")

    def avg(grp):
        vals = [it["chg_pct"] for it in items if it["group"] == grp]
        return round(sum(vals) / len(vals), 2) if vals else None

    korea_avg = avg("korea")
    asia_avg = avg("asia")   # 한국 제외 아시아 피어(대만·일본)
    us_chg = avg("us")   # 미국 반도체(마이크론·브로드컴·SOXX) 전일 평균
    asia_all = [it["chg_pct"] for it in items if it["group"] in ("korea", "asia")]
    strength = round(sum(asia_all) / len(asia_all), 2) if asia_all else None

    # 동조 판정 — 한국 반도체가 아시아 섹터와 함께 움직였나, 단독이었나
    regime, msg = "데이터 부족", ""
    if korea_avg is not None and asia_avg is not None:
        ks, asi = korea_avg, asia_avg
        if ks > 0.5 and asi > 0.5:
            regime = "🌏 아시아 동반 강세"
            msg = ("한국 반도체 강세는 아시아 섹터(대만·일본) 전반의 동조 흐름. "
                   "외국인 반도체 매수의 배경이며, 개장 전 미국 선물만으론 못 잡는 당일 동력.")
        elif ks > 0.5 and asi <= 0.5:
            regime = "🇰🇷 한국 단독 강세"
            msg = "아시아 피어는 잠잠한데 한국만 강세 — 한국 고유 수급. 되돌림 위험에 유의."
        elif ks < -0.5 and asi < -0.5:
            regime = "🌏 아시아 동반 약세"
            msg = "아시아 반도체 섹터 전반 약세 — 한국 약세는 글로벌 동조."
        elif ks < -0.5 and asi >= -0.5:
            regime = "🇰🇷 한국 단독 약세"
            msg = "아시아 피어는 버티는데 한국만 약세 — 한국 고유 악재 점검."
        else:
            regime = "⚖️ 중립·혼조"
            msg = "아시아 반도체 섹터 뚜렷한 방향 없음."

    print(f"\n  한국 {korea_avg}% · 아시아피어 {asia_avg}% · 미국 {us_chg}% → {regime}")

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "items": items,
        "korea_avg": korea_avg, "asia_peer_avg": asia_avg,
        "us_chg": us_chg, "strength": strength,
        "regime": regime, "message": msg,
        "note": ("KOSPI 장중(09:00~15:30 KST)과 겹치는 대만·일본 반도체를 당일 관측해 "
                 "'아시아 동반 랠리'인지 '한국 단독'인지 실시간 판별. 시세는 거래소·소스 지연(~15분) 가능."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
