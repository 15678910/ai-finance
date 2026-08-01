"""시장별 정규장 마감 기준 '미완성 봉' 판정 — 장중 수집 방어

문제
----
yfinance 일봉의 마지막 행은 해당 시장이 아직 개장 중이면 '종가 미확정' 상태다.
그 봉으로 이동평균·회귀·백테스트를 계산하면 **수집 시각에 따라 결과가 달라진다**.
(예: 13시에 돌린 추세 신호와 18시에 돌린 신호가 다름 → 재현 불가)

해결
----
마지막 봉이 '오늘 것'이고 그 시장이 아직 안 닫혔으면 해당 봉을 제외하고 계산한다.
제외 사실은 호출부가 사용자에게 표기할 수 있도록 날짜 문자열로 반환한다(은폐 금지).

적용 대상이 아닌 것
------------------
  · 선물(=F) · FX(=X) · 코인(-USD): 거의 24시간 연속거래라 '정규장 마감' 개념이 없음
  · 실시간 현재가를 의도적으로 보여주는 경로: 최신 값이 목적이므로 제외하면 안 됨

단위 테스트: python -m core.market_hours
"""

from datetime import datetime

try:
    import pandas as pd
except ImportError:                                   # pandas 없으면 판정 불가 → 항상 False
    pd = None

# 시장별 (IANA 타임존, 정규장 마감 시:분, 종가 확정 버퍼(분))
# 버퍼: 마감 직후에는 정산가가 확정 전이거나 데이터 제공사 반영이 늦을 수 있어 여유를 둠
MARKETS = {
    "KRX": ("Asia/Seoul",       (15, 30), 10),        # 한국거래소 정규장
    "JPX": ("Asia/Tokyo",       (15, 0),  10),        # 도쿄증권거래소 후장
    "US":  ("America/New_York", (16, 0),  15),        # 미국 정규장 (서머타임은 tz가 자동 처리)
}

CONTINUOUS_SUFFIX = ("=F", "=X")                      # 선물 · FX
KRX_INDEX = {"^KS11", "^KQ11", "^KS200"}
JPX_INDEX = {"^N225", "^TPX"}


def market_of(symbol):
    """티커 → 시장 코드. 연속거래 상품이면 None(판정 안 함)."""
    s = (symbol or "").strip().upper()
    if not s:
        return None
    if s.endswith(CONTINUOUS_SUFFIX) or s.endswith("-USD") or s.endswith("-KRW"):
        return None
    if s.endswith(".KS") or s.endswith(".KQ") or s in KRX_INDEX:
        return "KRX"
    if s.endswith(".T") or s in JPX_INDEX:
        return "JPX"
    return "US"                                       # 그 외는 미국 상장으로 간주


def _local_now(market, now=None):
    tzname = MARKETS[market][0]
    if now is None:
        return pd.Timestamp.now(tz=tzname)
    ts = pd.Timestamp(now)
    return ts.tz_localize("UTC").tz_convert(tzname) if ts.tz is None else ts.tz_convert(tzname)


def incomplete_last_bar(last_index_value, symbol=None, market=None, now=None):
    """마지막 봉이 '오늘 것 + 아직 장중'이면 True.

    last_index_value: DataFrame/Series 인덱스의 마지막 값 (tz-aware/naive 모두 허용).
                      yfinance 단일 티커는 거래소 tz, 멀티 티커는 tz-naive 거래소 날짜.
    """
    if pd is None or last_index_value is None:
        return False
    m = market or market_of(symbol)
    if m not in MARKETS:
        return False
    try:
        bar = pd.Timestamp(last_index_value)
        bar_date = (bar.tz_localize(None) if bar.tz is not None else bar).strftime("%Y-%m-%d")
        cur = _local_now(m, now)
    except Exception:
        return False
    if bar_date != cur.strftime("%Y-%m-%d"):
        return False
    (ch, cm), buf = MARKETS[m][1], MARKETS[m][2]
    return cur.hour * 60 + cur.minute < ch * 60 + cm + buf


def drop_incomplete(df, symbol=None, market=None, now=None):
    """(DataFrame|Series, 제외한 날짜 or None) 반환. 제외 대상이 없으면 원본 그대로."""
    if df is None or len(df) == 0:
        return df, None
    if incomplete_last_bar(df.index[-1], symbol, market, now):
        dropped = pd.Timestamp(df.index[-1])
        dropped = (dropped.tz_localize(None) if dropped.tz is not None else dropped).strftime("%Y-%m-%d")
        return df.iloc[:-1], dropped
    return df, None


def _selftest():
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    assert pd is not None, "pandas 필요"
    ok = 0

    def chk(cond, label):
        nonlocal ok
        assert cond, f"FAIL: {label}"
        ok += 1
        print(f"  ✓ {label}")

    chk(market_of("005930.KS") == "KRX", "005930.KS → KRX")
    chk(market_of("^KS11") == "KRX", "^KS11 → KRX")
    chk(market_of("^N225") == "JPX", "^N225 → JPX")
    chk(market_of("^SOX") == "US", "^SOX → US")
    chk(market_of("MU") == "US", "MU → US")
    chk(market_of("NQ=F") is None, "NQ=F → 판정 제외(선물)")
    chk(market_of("JPY=X") is None, "JPY=X → 판정 제외(FX)")

    # KRX: 마감 15:30 + 버퍼 10분 → 15:40 이전이면 미완성
    day = pd.Timestamp("2026-07-31")
    for hm, want in ((("13:17"), True), (("15:39"), True), (("15:40"), False), (("18:30"), False)):
        now = pd.Timestamp(f"2026-07-31 {hm}", tz="Asia/Seoul")
        chk(incomplete_last_bar(day, "005930.KS", now=now) is want,
            f"KRX 오늘봉 @ {hm} → 미완성={want}")
    chk(incomplete_last_bar(pd.Timestamp("2026-07-30"), "005930.KS",
                            now=pd.Timestamp("2026-07-31 13:17", tz="Asia/Seoul")) is False,
        "KRX 어제봉 @ 13:17 → 완성(제외 안 함)")

    # US: 마감 16:00 ET + 15분
    for hm, want in ((("11:00"), True), (("16:14"), True), (("16:15"), False)):
        now = pd.Timestamp(f"2026-07-31 {hm}", tz="America/New_York")
        chk(incomplete_last_bar(day, "^SOX", now=now) is want, f"US 오늘봉 @ {hm} ET → 미완성={want}")

    # JPX: 마감 15:00 JST + 10분
    chk(incomplete_last_bar(day, "^N225", now=pd.Timestamp("2026-07-31 15:09", tz="Asia/Tokyo")) is True,
        "JPX 오늘봉 @ 15:09 JST → 미완성")
    chk(incomplete_last_bar(day, "^N225", now=pd.Timestamp("2026-07-31 15:10", tz="Asia/Tokyo")) is False,
        "JPX 오늘봉 @ 15:10 JST → 완성")

    # tz-aware 인덱스(단일 티커 다운로드 형태)도 동일 판정
    chk(incomplete_last_bar(pd.Timestamp("2026-07-31", tz="Asia/Seoul"), "005930.KS",
                            now=pd.Timestamp("2026-07-31 13:17", tz="Asia/Seoul")) is True,
        "tz-aware 인덱스 @ 13:17 → 미완성")

    # drop_incomplete: 실제 프레임에서 마지막 행 제거
    idx = pd.to_datetime(["2026-07-29", "2026-07-30", "2026-07-31"])
    df = pd.DataFrame({"Close": [1.0, 2.0, 3.0]}, index=idx)
    d2, dropped = drop_incomplete(df, "005930.KS", now=pd.Timestamp("2026-07-31 13:17", tz="Asia/Seoul"))
    chk(len(d2) == 2 and dropped == "2026-07-31", "drop_incomplete 장중 → 마지막 행 제거 + 날짜 반환")
    d3, dropped3 = drop_incomplete(df, "005930.KS", now=pd.Timestamp("2026-07-31 18:30", tz="Asia/Seoul"))
    chk(len(d3) == 3 and dropped3 is None, "drop_incomplete 마감후 → 원본 유지")
    d4, dropped4 = drop_incomplete(df, "NQ=F", now=pd.Timestamp("2026-07-31 13:17", tz="Asia/Seoul"))
    chk(len(d4) == 3 and dropped4 is None, "선물은 제외 안 함")
    d5, _ = drop_incomplete(pd.DataFrame(), "005930.KS")
    chk(len(d5) == 0, "빈 프레임 방어")

    print(f"\n[OK] {ok}건 통과")


if __name__ == "__main__":
    _selftest()
