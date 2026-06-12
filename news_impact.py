"""
뉴스 시장영향 분류 (격화/완화)
================================
헤드라인을 지정학·시장 키워드로 격화(악재)/완화(호재)/중립 분류.
긴급뉴스 알림 태그 + 예측 반전경보(선물 vs 뉴스심리)에 공용 사용.
🚨 키워드 기반 근사치. 절대 단독 투자판단 금지.
"""

# 완화(호재) — 반등·우호
DEESC = [
    "deal", "truce", "ceasefire", "agreement", "agree", "resolve", "resolved",
    "ease", "easing", "de-escalat", "peace", "talks", "diplomacy", "rebound",
    "rally", "surge", "jump", "recover", "recovery", "soar", "gains", "rebounds",
    "완화", "휴전", "협상", "타결", "합의", "기대", "반등", "급등", "회복", "진정",
    "평화", "해소", "안도", "낙관",
]
# 격화(악재) — 하락·위험
ESC = [
    "strike", "attack", "war", "missile", "escalat", "selloff", "sell-off",
    "sell off", "plunge", "plummet", "crash", "sanction", "conflict", "invade",
    "invasion", "threat", "tension", "halt trading", "circuit breaker", "rout",
    "격화", "공격", "전쟁", "미사일", "긴장", "급락", "폭락", "제재", "충돌",
    "확전", "위기", "패닉", "서킷브레이커", "거래정지", "경고", "악화", "공포",
]
# 부정 패턴 — '완화 무산'(악재) / '공격 취소'(호재)
NEG_BAD = ["fade", "fail", "failed", "collapse", "reject", "breakdown",
           "off the table", "무산", "결렬", "거부", "철회 거부"]
NEG_GOOD = ["cancel", "called off", "call off", "averted", "avert", "취소",
            "철회", "중단"]


def classify_news(title):
    """헤드라인 → (score, label, emoji, impact)
    score: +1 완화 / -1 격화 / 0 중립·혼조."""
    t = (title or "").lower()
    has_d = any(k in t for k in DEESC)
    has_e = any(k in t for k in ESC)

    # 부정 패턴 우선 (예: 'deal hopes fade'=악재, 'strikes canceled'=호재)
    if has_d and any(k in t for k in NEG_BAD):
        return -1, "격화", "🔴", "추가 하락 위험"
    if has_e and any(k in t for k in NEG_GOOD):
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
