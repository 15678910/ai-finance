"""
뉴스 시장영향 분류 (격화/완화)
================================
헤드라인을 지정학·시장 키워드로 격화(악재)/완화(호재)/중립 분류.
긴급뉴스 알림 태그 + 예측 반전경보(선물 vs 뉴스심리)에 공용 사용.
🚨 키워드 기반 근사치. 절대 단독 투자판단 금지.

키워드 표기 규칙
---------------
  "war"    → 단어 단위로만 일치 (\\bwar\\b). warehouse·software 는 걸리지 않는다.
  "escalat*" → 어간. 뒤에 뭐가 붙어도 일치 (escalate·escalating·escalation).
  한글 키워드 → 부분 일치 (조사·어미가 붙는 교착어라 단어 경계를 쓸 수 없다.
                 '완화'가 '완화됐다'·'완화가' 안에 있어도 잡아야 한다).

왜 단어 경계인가 (2026-08 수정)
------------------------------
이전에는 전부 부분 일치라 오분류가 심했다. 실측 사례:
  · "Fed increases rates"  → increase ⊃ ease  → 완화(호재)로 오판
  · "software update"      → software ⊃ war   → 격화(하락 위험)
  · "an ideal entry point" → ideal ⊃ deal     → 완화
  · "route through"        → route ⊃ rout     → 격화
  · "disease outbreak"     → disease ⊃ ease   → 완화
  · "warns Iran against"   → against ⊃ gains  → 격화와 상쇄돼 '혼조'
이 오류는 뉴스 심리 지수와 예측 반전 경보까지 오염시켰다.
"""

import re

# 완화(호재) — 반등·우호
DEESC = [
    "deal", "deals", "truce", "ceasefire", "agreement", "agree*", "resolve*",
    "ease", "eases", "easing", "de-escalat*", "peace", "talks", "diplomacy",
    "rebound*", "rally", "rallies", "surge*", "jump*", "recover*", "soar*", "gains",
    "bullish",
    "완화", "휴전", "협상", "타결", "합의", "기대", "반등", "급등", "회복", "진정",
    "평화", "해소", "안도", "낙관",
]
# 격화(악재) — 하락·위험
ESC = [
    "strike", "strikes", "attack*", "war", "wars", "warn*", "missile*", "escalat*",
    "selloff", "sell-off", "sell off", "plunge*", "plummet*", "crash*", "sanction*",
    "conflict", "invade*", "invasion", "threat", "threats", "threaten*", "tension*",
    "halt trading", "circuit breaker", "rout",
    # 완화측 시세 동사(surge·jump·rally·soar)에 대응하는 하락 동사가 빠져 있어
    # 순심리가 상방으로 치우쳤다 — 대칭이 되도록 보강.
    "tumble*", "slump*", "bearish",
    "격화", "공격", "전쟁", "미사일", "긴장", "급락", "폭락", "제재", "충돌",
    "확전", "위기", "패닉", "서킷브레이커", "거래정지", "경고", "악화", "공포",
]
# 부정 패턴 — '완화 무산'(악재) / '공격 취소'(호재)
NEG_BAD = ["fade*", "fail*", "collapse*", "reject*", "breakdown",
           "off the table", "무산", "결렬", "거부", "철회 거부"]
NEG_GOOD = ["cancel*", "called off", "call off", "avert*", "취소",
            "철회", "중단"]


def _compile(keywords):
    """키워드 목록 → 하나의 정규식.

    ASCII 키워드는 단어 경계로 감싸고(어간은 뒤쪽 경계 생략),
    한글 등 비ASCII 키워드는 부분 일치로 둔다.
    """
    parts = []
    for kw in keywords:
        stem = kw.endswith("*")
        body = kw[:-1] if stem else kw
        esc = re.escape(body)
        if body.isascii():
            # 뒤 경계는 어간이면 생략 — escalat* 가 escalating 도 잡도록
            parts.append(r"\b" + esc + ("" if stem else r"\b"))
        else:
            parts.append(esc)          # 한글: 조사·어미 때문에 부분 일치 유지
    return re.compile("|".join(parts), re.IGNORECASE)


_RE_DEESC = _compile(DEESC)
_RE_ESC = _compile(ESC)
_RE_NEG_BAD = _compile(NEG_BAD)
_RE_NEG_GOOD = _compile(NEG_GOOD)


def classify_news(title):
    """헤드라인 → (score, label, emoji, impact)
    score: +1 완화 / -1 격화 / 0 중립·혼조."""
    t = (title or "").lower()
    has_d = bool(_RE_DEESC.search(t))
    has_e = bool(_RE_ESC.search(t))

    # 부정 패턴 우선 (예: 'deal hopes fade'=악재, 'strikes canceled'=호재)
    if has_d and _RE_NEG_BAD.search(t):
        return -1, "격화", "🔴", "추가 하락 위험"
    if has_e and _RE_NEG_GOOD.search(t):
        return +1, "완화", "🟢", "반등 우호"

    if has_d and not has_e:
        return +1, "완화", "🟢", "반등 우호"
    if has_e and not has_d:
        return -1, "격화", "🔴", "하락 위험"
    if has_d and has_e:
        return 0, "혼조", "⚪", "방향 혼조"
    return 0, "중립", "⚪", ""


def aggregate_sentiment(titles):
    """헤드라인 목록 → 순심리 점수·라벨. (esc_count, deesc_count, score, label, emoji)."""
    esc = deesc = 0
    for tt in titles:
        s, _, _, _ = classify_news(tt)
        if s > 0:
            deesc += 1
        elif s < 0:
            esc += 1
    score = deesc - esc
    if score >= 2:
        label, emoji = "우호 (완화 우세)", "🟢"
    elif score <= -2:
        label, emoji = "위험 (격화 우세)", "🔴"
    else:
        label, emoji = "중립", "⚪"
    return {"esc": esc, "deesc": deesc, "score": score, "label": label, "emoji": emoji}
