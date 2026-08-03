"""
공매도 잔고 모니터 — 관심종목 잔고비율·추세·신호등
====================================================
KRX는 2024년경부터 공매도 통계(정보데이터시스템·공매도종합포털 모두)에
회원 로그인을 요구 + GitHub 데이터센터 IP 차단 → GitHub cron 불가.
→ 내 PC(한국 IP) 로컬 스케줄러에서 pykrx(KRX 로그인)로 수집.
   KRX_ID/KRX_PW 는 Windows 사용자 환경변수(setx)로 설정 (코드/저장소에 미포함).

수집: 종목별 공매도 '잔고' 비율(시총 대비 %) 최근 ~20영업일 + 5/20일 추세
신호등(경험칙): <3% 정상 · 3~5% 주의 · ≥5% 경고/숏스퀴즈 후보
※ 잔고 데이터는 T+2 지연 공시(오늘 보는 값 = 2영업일 전 상황).
※ 대차잔고(KOFIA)는 별도 지표 — 이 파일은 '실행된 공매도 잔고'만 다룸.

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

        stocks.append({
            "code": code, "name": name, "asof": asof,
            "ratio": ratio, "chg_5d": chg5, "chg_20d": chg20,
            "trend": tr, "signal": sig, "signal_label": sig_label, "signal_color": sig_color,
            "balance_shares": bal,
            "spark": ratios[-20:],
        })
        print(f"  {name}: 잔고비율 {ratio}% ({asof}) 5일 {chg5:+} → {sig_label}")

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
