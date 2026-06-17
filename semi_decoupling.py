"""
반도체 디커플링 진단 — 한국 반도체(SK·삼성) vs 미국 SOX
==========================================================
"반도체 섹터가 미국과 따로 노는 날엔 예측이 구조적으로 빗나간다" → 그 디커플링을 측정한다.

  · 미국 SOX(필라델피아 반도체) 전일 등락 = 개장 전 알 수 있는 '반도체 선행신호'
  · 한국 반도체(D) ~ SOX(전일) 무절편 회귀 → 기대치 산출
  · 디커플링 = 실제 − 기대 = '한국 고유 반도체 성분'(장중 아시아 수급·AI테마·외인)
    → 이 부분은 개장 전 예측 불가. 진단·신뢰도 경고용(예측 아님).

예: 6/17 SOX −5.7%인데 SK +5.8% → 디커플링 +10.6%p (예측이 하락을 본 이유가 여기서 드러남).

출력: docs/semi_decoupling.json
"""

import json
import os
import re
import sys
import urllib.request
from datetime import datetime, date, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "semi_decoupling.json")
ASSETS = [("SK하이닉스", "000660"), ("삼성전자", "005930")]
W = 20


def naver_close(sym, days=120):
    end = date.today().strftime("%Y%m%d")
    start = (date.today() - timedelta(days=days)).strftime("%Y%m%d")
    url = (f"https://api.finance.naver.com/siseJson.naver?symbol={sym}"
           f"&requestType=1&startTime={start}&endTime={end}&timeframe=day")
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0", "Referer": "https://finance.naver.com/"})
    txt = urllib.request.urlopen(req, timeout=12).read().decode("utf-8")
    rows = re.findall(r'\["(\d{8})",\s*[\d.]+,\s*[\d.]+,\s*[\d.]+,\s*([\d.]+)', txt)
    return {f"{d[:4]}-{d[4:6]}-{d[6:]}": float(c) for d, c in rows}


def naver_foreign_flow(ticker):
    """frgn.naver 외국인 순매매(주식수)×종가 → 최근일·5일 순매수(억원). 로그인 불필요.
    디커플링의 실제 동력(외인 수급)을 측정."""
    try:
        url = f"https://finance.naver.com/item/frgn.naver?code={ticker}&page=1"
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0", "Referer": "https://finance.naver.com/"})
        html = urllib.request.urlopen(req, timeout=12).read().decode("euc-kr", errors="replace")
    except Exception:
        return None, None
    def pint(s):
        s = (s or "").replace(",", "").replace("+", "").strip()
        try:
            return int(float(s))
        except Exception:
            return 0
    out = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.DOTALL):
        if not re.search(r"\d{4}\.\d{2}\.\d{2}", tr):
            continue
        tds = [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", td)).strip()
               for td in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.DOTALL)]
        if len(tds) < 9:
            continue
        out.append((tds[0].replace(".", "-"), pint(tds[6]), pint(tds[1])))  # (date, 순매매주, 종가)
    out.sort(key=lambda x: x[0])
    if not out:
        return None, None
    f1 = round(out[-1][1] * out[-1][2] / 1e8)            # 최근일 억원
    f5 = round(sum(n * c for _, n, c in out[-5:]) / 1e8)  # 5일 합 억원
    return f1, f5


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    import numpy as np
    import yfinance as yf
    import warnings
    warnings.filterwarnings("ignore")

    sox = yf.download("^SOX", period="6mo", interval="1d", progress=False, auto_adjust=True)["Close"]
    if hasattr(sox, "columns"):
        sox = sox.iloc[:, 0]
    sox = sox.dropna()
    sr = {d.strftime("%Y-%m-%d"): float(v) for d, v in sox.pct_change().items() if not np.isnan(v)}
    sox_dates = sorted(sr.keys())
    sox_chg = round(sr[sox_dates[-1]] * 100, 2) if sox_dates else None
    sox_asof = sox_dates[-1] if sox_dates else None

    out_assets = []
    for nm, sym in ASSETS:
        try:
            px = naver_close(sym)
        except Exception as e:
            print(f"  [WARN] {nm} 수집 실패: {e}")
            continue
        kd = sorted(px)
        recs = []  # (D, 한국반도체ret, SOX전일ret)
        for i in range(1, len(kd)):
            D, P = kd[i], kd[i - 1]
            if P in sr and px[P] > 0:
                recs.append((D, px[D] / px[P] - 1, sr[P]))
        if len(recs) < W + 2:
            continue
        win = recs[-W:]
        Y = np.array([r[1] for r in win])
        X = np.array([r[2] for r in win])
        beta = float((X @ Y) / (float(X @ X) or 1e-9))
        r = float(np.corrcoef(X, Y)[0, 1])
        D, y, x = recs[-1]
        expected = beta * x
        resid = y - expected
        recent10 = float(np.mean([recs[j][1] - beta * recs[j][2] for j in range(len(recs) - 10, len(recs))]))
        if abs(resid) < 0.01:
            verdict, vkey = "미국 반도체와 동조", "sync"
        elif resid > 0:
            verdict, vkey = "한국 반도체 독자 강세", "kr_strong"
        else:
            verdict, vkey = "한국 반도체 독자 약세", "kr_weak"
        f1, f5 = naver_foreign_flow(sym)  # 외국인 순매수(억원) — 디커플링 동력
        out_assets.append({
            "name": nm, "code": sym, "date": D,
            "actual_pct": round(y * 100, 2), "expected_pct": round(expected * 100, 2),
            "decoupling_pp": round(resid * 100, 2), "beta": round(beta, 2), "r": round(r, 2),
            "recent10_pp": round(recent10 * 100, 2), "verdict": verdict, "vkey": vkey,
            "foreign_1d_eok": f1, "foreign_5d_eok": f5,
        })
        print(f"  {nm} {D}: 실제 {y*100:+.2f}% vs SOX기대 {expected*100:+.2f}% "
              f"→ 디커플링 {resid*100:+.2f}%p (β{beta:.2f} r{r:+.2f}) {verdict} | 외인5일 {f5}억")

    # 레짐(SK 기준): 최근 디커플링 크기·지속
    sk = next((a for a in out_assets if a["code"] == "000660"), None)
    if sk and abs(sk["decoupling_pp"]) >= 3:
        regime = {"key": "high", "text": "🔴 반도체 디커플링 큼 — 개장 전 선물·미 반도체 신호 신뢰도 낮음(한국 고유 수급 지배). 예측 빗나갈 위험."}
    elif sk and abs(sk["recent10_pp"]) >= 1:
        regime = {"key": "mid", "text": "🟡 반도체 디커플링 진행 — 한국 반도체가 미국과 다소 따로 움직이는 국면. 예측 신뢰도 주의."}
    else:
        regime = {"key": "low", "text": "🟢 미국 반도체와 대체로 동조 — 선물·SOX 선행신호 신뢰도 상대적 양호."}

    # 외인 수급 편향(약한 선행 r~0.22 — 예측 입력 아닌 신뢰도 편향용)
    sk_f5 = sk.get("foreign_5d_eok") if sk else None
    if sk_f5 is not None and abs(sk_f5) >= 10000:
        regime["foreign_bias"] = ("외인 강매수 지속 → 상방 디커플링 편향(선물 하락예측 신뢰 ↓)"
                                  if sk_f5 > 0 else "외인 강매도 지속 → 하방 위험(반도체 적신호)")

    # 외국인 반도체 수급 1일 급변(±3000억↑) 텔레그램 경보(하루 1회 dedup)
    surge = [a for a in out_assets if a.get("foreign_1d_eok") and abs(a["foreign_1d_eok"]) >= 3000]
    if surge:
        try:
            from core import send_message, get_secret, load_state, save_state
            today_s = datetime.now(KST).strftime("%Y-%m-%d")
            st = load_state("semi_foreign_alert", default={})
            if get_secret("TELEGRAM_FINANCE_BOT_TOKEN") and st.get("last") != today_s:
                lines = [f"{a['name']} 외인 {a['foreign_1d_eok']:+,}억 — "
                         f"{'강매수(상방 디커플링)' if a['foreign_1d_eok'] > 0 else '강매도(하방 위험)'}"
                         for a in surge]
                if send_message("💰 외국인 반도체 수급 급변\n" + "\n".join(lines)
                                + "\n※ 디커플링 동력(약한 선행 r~0.22). 선물 예측과 함께 보세요."):
                    save_state("semi_foreign_alert", {"last": today_s})
                    print("  💰 외인 수급 급변 텔레그램 발송")
        except Exception as e:
            print(f"  [WARN] 외인 경보 실패: {e}")

    out = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "window": W,
        "sox": {"chg_pct": sox_chg, "asof": sox_asof},
        "assets": out_assets,
        "regime": regime,
        "note": ("미국 SOX 전일 등락은 개장 전 확인 가능(반도체 선행신호). 한국 고유 디커플링(실제−SOX기대)의 "
                 "실제 동력은 아래 '외국인 5일 순매수' — 외인이 크게 사면 미국과 따로 강세(6/17 SK처럼). "
                 "이 수급은 장중 결정이라 개장 전 예측 불가 → 진단·신뢰도 경고용. 동시상관, 예측 아님."),
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"  미국 SOX 최근: {sox_chg}% · 레짐 {regime['key']}")
    print(f"[OK] {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
