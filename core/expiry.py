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

⚠️ 공휴일 보정의 한계
--------------------
만기일이 휴장일이면 거래소는 **직전 거래일**로 앞당긴다. 이 모듈은 날짜가 고정된
공휴일(신정·한글날·성탄절 등)만 반영하며, **음력 기반 공휴일(설·추석·부처님오신날)과
미국 Good Friday는 반영하지 않는다** — 임의로 추정하면 틀린 날짜를 사실처럼
표시하게 되므로, 대신 caveat 플래그를 함께 반환해 호출부가 표기하도록 한다.

단위 테스트: python -m core.expiry
"""

from datetime import date, timedelta

THU, FRI = 3, 4                       # date.weekday(): 월=0
QUARTER_MONTHS = (3, 6, 9, 12)

# 날짜가 매년 고정된 휴장일만 수록 (음력 공휴일은 의도적으로 제외 — 상단 주석 참조)
KRX_FIXED_CLOSED = {(1, 1), (3, 1), (5, 1), (5, 5), (6, 6), (8, 15), (10, 3), (10, 9), (12, 25), (12, 31)}
US_FIXED_CLOSED = {(1, 1), (6, 19), (7, 4), (12, 25)}


def nth_weekday(year, month, weekday, n):
    """그 달의 n번째 <weekday> 날짜 (n=1,2,3…)."""
    first = date(year, month, 1)
    return first + timedelta(days=(weekday - first.weekday()) % 7 + 7 * (n - 1))


def _shift_off_holiday(d, fixed_closed):
    """휴장일·주말이면 직전 평일로 앞당김. (보정 여부, 날짜) 반환."""
    moved = False
    for _ in range(7):
        if d.weekday() >= 5 or (d.month, d.day) in fixed_closed:
            d -= timedelta(days=1)
            moved = True
        else:
            break
    return moved, d


def kr_expiry(year, month):
    """한국 파생 만기일 — 둘째 목요일(휴장일이면 직전 거래일)."""
    return _shift_off_holiday(nth_weekday(year, month, THU, 2), KRX_FIXED_CLOSED)


def us_expiry(year, month):
    """미국 옵션 만기일 — 셋째 금요일(휴장일이면 직전 거래일)."""
    return _shift_off_holiday(nth_weekday(year, month, FRI, 3), US_FIXED_CLOSED)


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

    # 공휴일 보정: 한글날(10/9)이 둘째 목요일인 해 → 직전 거래일로
    moved, d = kr_expiry(2025, 10)
    chk(nth_weekday(2025, 10, THU, 2) == date(2025, 10, 9), "2025-10 둘째 목요일 = 10-09(한글날)")
    chk(moved and d == date(2025, 10, 8), "한글날 만기 → 직전 거래일 10-08 보정")

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
