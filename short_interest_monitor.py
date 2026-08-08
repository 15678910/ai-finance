"""
공매도 잔고 모니터 — 관심종목 잔고비율·추세·신호등
====================================================
KRX는 2024년경부터 공매도 통계(정보데이터시스템·공매도종합포털 모두)에
회원 로그인을 요구 + GitHub 데이터센터 IP 차단 → GitHub cron 불가.
→ 내 PC(한국 IP) 로컬 스케줄러에서 pykrx(KRX 로그인)로 수집.
   KRX_ID/KRX_PW 는 Windows 사용자 환경변수(setx)로 설정 (코드/저장소에 미포함).

수집 — 성격이 다른 두 지표를 함께 담는다
  1) 공매도 '잔고' 비율 (시총 대비 %)  · T+2 · 누적 잔량 → 느린 배경 지표
  2) 공매도 '거래' 비중 (거래량 대비 %) · T+1 · 하루치 흐름 → 하루 빠른 민감 지표

  둘 다 '비중'이라 불리지만 분모가 다르다. 섞어 읽으면 안 된다.
    잔고비중 = 공매도 잔고주식수 ÷ 상장주식수
    거래비중 = 그날 공매도 거래량 ÷ 그날 전체 거래량
  잔고가 정체여도 거래가 튀는 날이 있다 — 그 움직임은 잔고만 봐서는 안 보인다.
  (실측 2026-08-07 한미반도체: 잔고 5일 +0.02%p 정체인데 거래비중 18.6%→28.6%)

신호등(경험칙): 잔고 <3% 정상 · 3~5% 주의 · ≥5% 경고/숏스퀴즈 후보
※ 어느 쪽도 실시간이 아니다. 거래는 다음 영업일, 잔고는 2영업일 뒤 공표된다.
※ 대차잔고(KOFIA)는 또 다른 지표 — 이 파일은 '실행된 공매도'만 다룬다.

출력: docs/short_interest.json
🚨 참고용 · 투자자문 아님. 공매도 잔고는 단독 지표로 쓰지 말 것(펀더멘탈·수급과 결합).
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "short_interest.json")

# 감시 종목 (필요 시 여기에 추가)
WATCH = [
    ("000660", "SK하이닉스"),
    ("005930", "삼성전자"),
    ("108490", "로보티즈"),
    ("042700", "한미반도체"),
    ("009150", "삼성전기"),
    ("373220", "LG에너지솔루션"),
    ("003550", "LG"),
    ("066570", "LG전자"),
]


def classify(ratio):
    """경험칙 신호등: <3 정상 / 3~5 주의 / >=5 경고(스퀴즈 후보)."""
    if ratio is None:
        return "unknown", "—", "gray"
    if ratio >= 5:
        return "warning", "경고·스퀴즈 후보", "red"
    if ratio >= 3:
        return "caution", "주의", "yellow"
    return "normal", "정상", "green"


def trend_of(vals):
    """최근 5영업일 잔고비율 변화 → up/down/flat."""
    if len(vals) < 6:
        return "flat", 0.0
    chg = vals[-1] - vals[-6]
    if chg > 0.15:
        return "up", round(chg, 2)
    if chg < -0.15:
        return "down", round(chg, 2)
    return "flat", round(chg, 2)


def fetch_volume(stock_mod, code, frm, to):
    """공매도 '거래' 비중 (T+1) — 잔고보다 하루 빠르고 하루치 움직임에 민감하다.

    반환: {asof, ratio, avg_5d, vs_avg_pp, shares, spike} 또는 None
      ratio  = 그날 공매도 거래량 ÷ 전체 거래량 × 100
      spike  = 최근 5일 평균보다 1.5배 이상 (갑자기 몰린 날)
    실패해도 잔고 수집은 계속되어야 하므로 예외를 삼키고 None 을 돌려준다.
    """
    try:
        df = stock_mod.get_shorting_volume_by_date(frm, to, code)
    except Exception as e:
        print(f"    [WARN] 거래 비중 조회 실패: {str(e)[:60]}")
        return None
    if df is None or len(df) == 0:
        return None
    rcol = next((c for c in df.columns if "비중" in str(c)), None)
    scol = next((c for c in df.columns if "공매도" in str(c)), None)
    if rcol is None:
        return None
    vals = [round(float(v), 2) for v in df[rcol].tolist()]
    last = df.index[-1]
    asof = last.strftime("%Y-%m-%d") if hasattr(last, "strftime") else str(last)[:10]
    recent = vals[-6:-1] if len(vals) >= 6 else vals[:-1]
    avg5 = round(sum(recent) / len(recent), 2) if recent else None
    out = {"asof": asof, "ratio": vals[-1], "avg_5d": avg5,
           "vs_avg_pp": round(vals[-1] - avg5, 2) if avg5 is not None else None,
           "spike": bool(avg5 and vals[-1] >= avg5 * 1.5),
           "series": vals[-20:]}
    if scol is not None:
        try:
            out["shares"] = int(df[scol].iloc[-1])
        except Exception:
            pass
    return out


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    if not (os.environ.get("KRX_ID") and os.environ.get("KRX_PW")):
        print("[INFO] KRX_ID/KRX_PW 환경변수 미설정 — 공매도 잔고는 KRX 로그인 필수라 수집 불가.")
        print("       setx KRX_ID / setx KRX_PW 로 설정 후 재실행하세요. 기존 파일 보존.")
        return 0                       # 스케줄러 전체 실패 방지

    try:
        from pykrx import stock
    except Exception as e:
        print(f"[ERROR] pykrx import 실패: {e}")
        return 1

    # 로그인 사전 점검 — 실패 시 즉시 중단(종목마다 재시도해 계정 잠금되는 것 방지)
    try:
        from pykrx.website.comm.auth import login_krx
        if not login_krx(os.environ["KRX_ID"], os.environ["KRX_PW"]):
            print("[ERROR] KRX 로그인 거부 — ID/비밀번호를 확인하세요. (반복 시도 방지 위해 중단, 기존 파일 보존)")
            return 1
    except ImportError:
        pass                            # pykrx 구조 변경 시 사전점검 생략

    now = datetime.now(KST)
    frm = (now - timedelta(days=45)).strftime("%Y%m%d")   # ~20영업일 확보
    to = now.strftime("%Y%m%d")

    stocks, asof_seen = [], []
    for code, name in WATCH:
        try:
            df = stock.get_shorting_balance_by_date(frm, to, code)
        except Exception as e:
            print(f"  [WARN] {name}({code}) 조회 실패: {e}")
            continue
        if df is None or len(df) == 0:
            print(f"  [WARN] {name}({code}) 데이터 없음")
            continue

        # 비중(잔고비율 %) 컬럼 탐색 — pykrx 버전에 따라 '비중' 명칭
        rcol = next((c for c in df.columns if "비중" in str(c)), None)
        bcol = next((c for c in df.columns if "잔고" in str(c) and "금액" not in str(c)), None)
        if rcol is None:
            print(f"  [WARN] {name}({code}) 비중 컬럼 없음: {list(df.columns)}")
            continue

        ratios = [round(float(v), 2) for v in df[rcol].tolist()]
        last = df.index[-1]
        asof = last.strftime("%Y-%m-%d") if hasattr(last, "strftime") else str(last)[:10]
        asof_seen.append(asof)
        ratio = ratios[-1]
        tr, chg5 = trend_of(ratios)
        chg20 = round(ratios[-1] - ratios[0], 2) if len(ratios) >= 15 else None
        sig, sig_label, sig_color = classify(ratio)
        bal = None
        if bcol is not None:
            try:
                bal = int(df[bcol].iloc[-1])
            except Exception:
                pass

        vol = fetch_volume(stock, code, frm, to)      # 거래 비중(T+1) — 잔고보다 하루 빠름
        stocks.append({
            "code": code, "name": name, "asof": asof,
            "ratio": ratio, "chg_5d": chg5, "chg_20d": chg20,
            "trend": tr, "signal": sig, "signal_label": sig_label, "signal_color": sig_color,
            "volume": vol,
            "balance_shares": bal,
            "spark": ratios[-20:],
        })
        print(f"  {name}: 잔고비율 {ratio}% ({asof}) 5일 {chg5:+} → {sig_label}")
        if vol:
            mark = " ⚡급증" if vol.get("spike") else ""
            va = f" (5일평균 {vol['avg_5d']}% 대비 {vol['vs_avg_pp']:+}%p)" if vol.get("avg_5d") is not None else ""
            print(f"      거래비중 {vol['ratio']}% ({vol['asof']}){va}{mark}")

    if not stocks:
        print("[ERROR] 전 종목 수집 실패 — 기존 파일 보존.")
        return 1

    out = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
        "asof": max(asof_seen) if asof_seen else None,
        "stocks": stocks,
        "thresholds": {"caution": 3, "warning": 5},
        "note": ("KRX 공매도 '잔고' 비율(시가총액 대비 %, T+2 지연 공시). "
                 "경험칙: <3% 정상 · 3~5% 주의 · ≥5% 경고/숏스퀴즈 후보 — 공식 기준 아님. "
                 "잔고 추세(증감)가 절대치보다 중요. 대차잔고(KOFIA)는 별도 지표(대기 수요). "
                 "펀더멘탈·외국인 수급과 반드시 결합 해석. 투자자문 아님."),
    }
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"[OK] {OUTPUT_FILE}  ({len(stocks)}종목 · 기준 {out['asof']})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
