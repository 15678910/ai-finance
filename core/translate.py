"""
영어 → 한국어 자동 번역 헬퍼
============================

Google Translate 무료(unauthenticated) endpoint 사용.
인증 키 불필요. 실패 시 None 반환 (호출부에서 원문 폴백).

여러 모듈(clarity_act_monitor, semi_challenger_monitor 등)에서 공유.
"""

import json
import urllib.request
import urllib.parse


def translate_to_korean(text: str, max_chars: int = 1500, timeout: int = 10) -> str | None:
    """영어 → 한국어 번역.

    응답 형식: [[["번역결과", "원문", null, null, ...], ...], ...]
    """
    if not text or not text.strip():
        return None
    text = text.strip()
    if len(text) > max_chars:
        text = text[:max_chars]

    try:
        params = {
            "client": "gtx",
            "sl": "en",
            "tl": "ko",
            "dt": "t",
            "q": text,
        }
        url = "https://translate.googleapis.com/translate_a/single?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (compatible; ai-finance-translator/1.0)",
            "Accept": "application/json",
        })
        with urllib.request.urlopen(req, timeout=timeout) as resp:  # nosec — 공개 endpoint
            raw = resp.read()
        data = json.loads(raw.decode("utf-8"))
        if not isinstance(data, list) or not data or not isinstance(data[0], list):
            return None
        translated_parts = []
        for seg in data[0]:
            if isinstance(seg, list) and seg and isinstance(seg[0], str):
                translated_parts.append(seg[0])
        result = "".join(translated_parts).strip()
        return result if result else None
    except Exception:
        # 무료 endpoint는 가끔 429/503 — 조용히 실패하고 원문 사용
        return None
