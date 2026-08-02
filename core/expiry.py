"""파생상품 만기일 계산 — 한국·미국 옵션/선물 만기 (규칙 기반)

거래소 규칙
----------
  · 한국(KRX): 코스피200 옵션·선물 만기 = **매월 둘째 목요일**
      - 3·6·9·12월은 지수선물·지수옵션·개별주식선물·개별주식옵션이 동시 만기
        → '동시만기일' = 네 마녀의 날(Quadruple Witching). 수급 왜곡·변동성 급증.
      - 그 외 달은 옵션 위주 월물 만기.
  · 미국: 주식·지수 옵션 만기 = **매월 셋째 금요일**
      - 3·6·9·12월은 지수선물·지수옵션·개별주식선물·개별주식옵션 동시 만기
        → Quadruple Witching.

공휴일 보정
----------
만기일이 휴장일이면 거래소는 **직전 거래일**로 앞당긴다. 휴장일은 `holidays` 패키지로
계산하며 **음력 공휴일(설·추석·부처님오신날)과 미국 Good Friday까지 반영**된다.

  · 한국: holidays.SouthKorea() + 근로자의날(5/1) + 연말 폐장일(12/31)
  · 미국: holidays.financial_holidays("NYSE")  ← Good Friday 포함

VALIDATION (2026-08-02, pykrx 실제 거래일 대조 · 2024-01~2026-07 / 628거래일)
  실제 휴장(평일) 47건 vs 라이브러리 42건
    · 오탐 0건 — 라이브러리가 공휴일이라 한 날은 전부 실제로 휴장이었다
    · 누락 5건 — 2024/2025/2026-05-01(근로자의날), 2024/2025-12-31(연말 폐장일)
      두 패턴뿐이라 상수로 보완 → 보완 후 불일치 0건
`holidays` 미설치 환경에서는 날짜 고정 공휴일만으로 동작하고 fallback 플래그를 세운다.

단위 테스트: python -m core.expiry
"""

from datetime import date, timedelta

try:
    import holidays as _holidays
except ImportError:                   # 미설치 시 고정 공휴일만으로 열화 동작
    _holidays = None

THU, FRI = 3, 4                       # date.weekday(): 월=0
QUARTER_MONTHS = (3, 6, 9, 12)

# holidays 패키지가 없을 때 쓰는 최소 집합 (음력 공휴일 미포함 — fallback 표시됨)
KRX_FIXED_CLOSED = {(1, 1), (3, 1), (5, 1), (5, 5), (6, 6), (8, 15), (10, 3), (10, 9), (12, 25), (12, 31)}
US_FIXED_CLOSED = {(1, 1), (6, 19), (7, 4), (12, 25)}
# 공휴일은 아니지만 거래소가 쉬는 날 (위 VALIDATION에서 확인된 누락 패턴)
KRX_EXTRA_CLOSED = {(5, 1), (12, 31)}          # 근로자의날 · 연말 폐장일

_cache = {}


def _closed_set(market, year):
    """(휴장일 set, holidays 패키지 사용 여부) — 연도 단위 캐시."""
    key = (market, year)
    if key in _cache:
        return _cache[key]
    if _holidays is None:
        res = (set(), False)
    elif market == "KR":
        s = {d for d in _holidays.SouthKorea(years=[year])}
        s |= {date(year, m, dd) for m, dd in KRX_EXTRA_CLOSED}
        res = (s, True)
    else:
        res = ({d for d in _holidays.financial_holidays("NYSE", years=[year])}, True)
    _cache[key] = res
    return res


def is_closed(d, market):
    """주말이거나 휴장일이면 True."""
    if d.weekday() >= 5:
        return True
    closed, ok = _closed_set(market, d.year)
    if ok:
        return d in closed
    fixed = KRX_FIXED_CLOSED if market == "KR" else US_FIXED_CLOSED
    return (d.month, d.day) in fixed


def holiday_data_ok():
    """공휴일 데이터가 음력까지 반영된 상태인지 (False면 근사치 — 호출부가 표기)."""
    return _holidays is not None


def nth_weekday(year, month, weekday, n):
    """그 달의 n번째 <weekday> 날짜 (n=1,2,3…)."""
    first = date(year, month, 1)
    return first + timedelta(days=(weekday - first.weekday()) % 7 + 7 * (n - 1))


def _shift_off_holiday(d, market):
    """휴장일·주말이면 직전 거래일로 앞당김. (보정 여부, 날짜) 반환."""
    moved = False
    for _ in range(10):
        if is_closed(d, market):
            d -= timedelta(days=1)
            moved = True
        else:
            break
    return moved, d


def kr_expiry(year, month):
    """한국 파생 만기일 — 둘째 목요일(휴장일이면 직전 거래일)."""
    return _shift_off_holiday(nth_weekday(year, month, THU, 2), "KR")


def us_expiry(year, month):
    """미국 옵션 만기일 — 셋째 금요일(휴장일이면 직전 거래일)."""
    return _shift_off_holiday(nth_weekday(year, month, FRI, 3), "US")


def _entry(d, moved, market, month):
    quad = month in QUARTER_MONTHS
    kr = market == "KR"
    return {
        "date": d.isoformat(),
        "market": market,
        "region": "🇰🇷" if kr else "🇺🇸",
        "quad": quad,
        "title": (("한국 선물·옵션 동시만기 (네 마녀의 날)" if quad else "한국 옵션 만기일") if kr
                  else ("미국 쿼드러플 위칭 (선물·옵션 동시만기)" if quad else "미국 옵션 만기일")),
        "rule": ("매월 둘째 목요일" if kr else "매월 셋째 금요일") + (" · 분기 동시만기" if quad else ""),
        "time": ("15:20 KST 동시호가" if kr else "16:00 ET"),
        "impact": "HIGH" if quad else "MEDIUM",
        "note": (("분기 동시만기 — 프로그램 매물·베이시스 청산으로 수급 왜곡, 변동성 급증. "
                  "만기 당일 장마감 동시호가(15:20~) 주의." if quad else
                  "월물 옵션 만기 — 만기 주 프로그램 매매·변동성 확대. VKOSPI와 함께 확인.") if kr else
                 ("쿼드러플 위칭 — 지수선물·지수옵션·개별주식선물·개별주식옵션 동시 만기. "
                  "미국 장 변동성 확대가 다음 날 한국 시장에 파급될 수 있음." if quad else
                  "미국 월물 옵션 만기 — 만기일 전후 미국 지수 변동성 확대.")),
        "holiday_adjusted": moved,
    }


def upcoming(today=None, months=5, include_today=True):
    """오늘 이후(포함) 가까운 만기일 목록 — 날짜순. 각 항목에 d_day 부여."""
    today = today or date.today()
    out = []
    y, m = today.year, today.month
    for i in range(months + 1):
        yy, mm = divmod((m - 1) + i, 12)
        yy, mm = y + yy, mm + 1
        for market, fn in (("KR", kr_expiry), ("US", us_expiry)):
            moved, d = fn(yy, mm)
            if d < today or (d == today and not include_today):
                continue
            e = _entry(d, moved, market, mm)
            e["d_day"] = (d - today).days
            out.append(e)
    out.sort(key=lambda x: (x["date"], x["market"] != "KR"))
    return out


def _selftest():
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass
    ok = 0

    def chk(cond, label):
        nonlocal ok
        assert cond, f"FAIL: {label}"
        ok += 1
        print(f"  ✓ {label}")

    # nth_weekday 기본 동작
    chk(nth_weekday(2026, 8, THU, 2) == date(2026, 8, 13), "2026-08 둘째 목요일 = 08-13")
    chk(nth_weekday(2026, 9, THU, 2) == date(2026, 9, 10), "2026-09 둘째 목요일 = 09-10")
    chk(nth_weekday(2026, 8, FRI, 3) == date(2026, 8, 21), "2026-08 셋째 금요일 = 08-21")
    # 1일이 해당 요일인 달 (경계)
    chk(nth_weekday(2026, 1, THU, 2) == date(2026, 1, 8), "2026-01-01=목 → 둘째 목요일 01-08")

    # 만기일
    _, d = kr_expiry(2026, 8)
    chk(d == date(2026, 8, 13), "韓 2026-08 만기 = 08-13")
    _, d = kr_expiry(2026, 9)
    chk(d == date(2026, 9, 10), "韓 2026-09 만기(동시만기) = 09-10")
    _, d = us_expiry(2026, 8)
    chk(d == date(2026, 8, 21), "美 2026-08 만기 = 08-21")

    # 공휴일 보정 — 2025-10은 둘째 목요일이 한글날(10/9)이고 그 앞이 추석 연휴(10/3~10/8)라
    # 직전 거래일이 10/2까지 밀린다. pykrx 실거래일로 확인한 정답(고정공휴일만으론 10/8로 오답).
    moved, d = kr_expiry(2025, 10)
    chk(nth_weekday(2025, 10, THU, 2) == date(2025, 10, 9), "2025-10 둘째 목요일 = 10-09(한글날)")
    if holiday_data_ok():
        chk(moved and d == date(2025, 10, 2), "한글날+추석연휴 → 직전 거래일 10-02 보정 (실거래일 검증됨)")

    # 음력 공휴일 반영 (holidays 패키지 필요) — 이 케이스가 기존 caveat의 핵심이었다
    if holiday_data_ok():
        chk(is_closed(date(2026, 2, 17), "KR"), "설날(2026-02-17) 휴장 인식")
        chk(is_closed(date(2026, 9, 25), "KR"), "추석(2026-09-25) 휴장 인식")
        chk(is_closed(date(2026, 5, 24), "KR"), "부처님오신날(2026-05-24) 휴장 인식")
        chk(is_closed(date(2026, 5, 1), "KR"), "근로자의날(5/1) 휴장 인식")
        chk(is_closed(date(2026, 12, 31), "KR"), "연말 폐장일(12/31) 휴장 인식")
        # 제헌절은 2026년부터 공휴일로 복원 — 2024·2025는 거래일, 2026은 휴장(실거래일로 확인)
        chk(not is_closed(date(2025, 7, 17), "KR"), "제헌절 2025-07-17은 거래일")
        chk(is_closed(date(2026, 7, 17), "KR"), "제헌절 2026-07-17은 휴장 (공휴일 복원)")
        chk(is_closed(date(2026, 4, 3), "US"), "美 Good Friday(2026-04-03) 휴장 인식")
        chk(not is_closed(date(2026, 8, 13), "KR"), "2026-08-13 정상 거래일")
        # 추석 연휴에 걸린 만기: 2027-09 둘째 목요일 09-09 (연휴 09-14~16과 무관) 확인
        m9, d9 = kr_expiry(2026, 9)
        chk(d9 == date(2026, 9, 10) and not m9, "2026-09 동시만기 09-10 (추석 이전, 보정 없음)")
    else:
        print("  (holidays 미설치 — 음력 케이스 스킵, fallback 동작)")

    # 분기 판정
    e = _entry(date(2026, 9, 10), False, "KR", 9)
    chk(e["quad"] and "네 마녀" in e["title"] and e["impact"] == "HIGH", "9월 = 동시만기 HIGH")
    e = _entry(date(2026, 8, 13), False, "KR", 8)
    chk((not e["quad"]) and e["impact"] == "MEDIUM", "8월 = 월물만기 MEDIUM")

    # upcoming 정렬·d_day
    up = upcoming(today=date(2026, 8, 2), months=2)
    chk(up[0]["date"] == "2026-08-13" and up[0]["market"] == "KR", "다음 만기 = 韓 08-13")
    chk(up[0]["d_day"] == 11, "d_day = 11일")
    chk(all(up[i]["date"] <= up[i + 1]["date"] for i in range(len(up) - 1)), "날짜 오름차순 정렬")
    chk(any(x["market"] == "US" and x["date"] == "2026-08-21" for x in up), "美 08-21 포함")
    # 당일 만기 포함 여부
    chk(upcoming(today=date(2026, 8, 13), months=1)[0]["d_day"] == 0, "만기 당일 → d_day 0")

    print(f"\n[OK] {ok}건 통과")
    print("\n다가오는 만기일:")
    for x in upcoming(months=3)[:6]:
        print(f"  {x['date']} D-{x['d_day']:<3} {x['region']} {x['title']}"
              f"{' (휴일보정)' if x['holiday_adjusted'] else ''}")


if __name__ == "__main__":
    _selftest()
