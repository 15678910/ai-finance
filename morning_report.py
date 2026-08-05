"""
모닝 브리핑 — 대시보드 수집 데이터 종합분석 보고서 (매일 아침 텔레그램 발송)
==================================================================================
docs/*.json (밤사이 해외·예측·심리·추세신호·공매도·개인수급·밸류·일정·뉴스)을
규칙 기반으로 종합해 한 편의 아침 보고서로 구성, 텔레그램으로 발송한다.
LLM 미사용(무료·결정적). 각 데이터는 없으면 해당 섹션 생략(방어적).

실행: GitHub Actions cron 21:15 UTC (= KST 06:15 예약, 큐 지연으로 실제 06:15~07:15 도착).
수동: python morning_report.py
🚨 정보 요약 · 투자자문 아님.
"""

import json
import os
import sys
from datetime import date, datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE_DIR, "docs")
SITE = "https://15678910.github.io/ai-finance/"
WD = ["월", "화", "수", "목", "금", "토", "일"]


def L(name):
    try:
        with open(os.path.join(DOCS, f"{name}.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def pct(v):
    try:
        return f"{float(v):+.2f}%"
    except Exception:
        return "—"


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
        except Exception:
            pass

    now = datetime.now(KST)
    S = []                                              # 섹션 리스트 (섹션별 문자열)

    # 생성 시각은 실제 값을 쓴다 — GitHub Actions 예약 실행은 상습 지연되므로
    # '07:00'을 하드코딩하면 07:50에 만든 리포트가 07:00짜리로 보인다(데이터 기준 시점 오해).
    S.append(f"🌅 AI Finance 모닝 브리핑\n{now:%Y-%m-%d} ({WD[now.weekday()]}) {now:%H:%M} KST 생성")

    # ── 밤사이 해외 ──
    ov = L("overseas_market")
    mk = (ov.get("markets") or {})
    lines = []
    # 앞에서부터 [:4]로 자르면 목록 뒤쪽의 '필라델피아 반도체'가 매번 잘려나갔다.
    # 반도체 비중이 큰 관심종목 구성상 SOX는 나스닥보다도 중요한 선행지표라 우선순위를 명시한다.
    PRIO = ["필라델피아 반도체", "Nasdaq", "S&P 500", "Dow Jones", "Russell 2000", "VIX",
            "Nasdaq 선물", "S&P 선물", "Dow 선물"]
    pool = [it for grp in ("미국_지수", "미국_선물") for it in (mk.get(grp) or [])]
    pool.sort(key=lambda it: PRIO.index(it.get("name")) if it.get("name") in PRIO else len(PRIO))
    for it in pool:
        nm = it.get("name") or it.get("label") or "?"
        ch = it.get("change_pct", it.get("chg_pct", it.get("chg")))
        if ch is not None:
            arrow = "🔺" if float(ch) >= 0 else "🔻"
            tag = " ⭐" if nm == "필라델피아 반도체" else ""      # 국내 반도체 직결 지표
            lines.append(f"  {arrow} {nm} {pct(ch)}{tag}")
    if lines:
        S.append("🌎 밤사이 해외\n" + "\n".join(lines[:9]))

    # ── 오늘 시초가·종합신호 ──
    cs = L("composite_signal")
    kp = ov.get("kospi_open_prediction")
    lines = []
    if kp:
        lines.append(f"  시초가: {kp}")
    if cs.get("bias"):
        fut = f" (야간선물 {pct(cs.get('futures_pct'))})" if cs.get("futures_pct") is not None else ""
        lines.append(f"  종합신호: {cs['bias']} · 점수 {cs.get('score')}{fut}")
    if cs.get("decoupling_warn"):
        lines.append(f"  ⚠️ {cs.get('warn_msg') or '디커플링 경고 — 선물신호 신뢰도 낮음'}")
    if lines:
        S.append("🔮 오늘 KOSPI\n" + "\n".join(lines))

    # ── 시장 심리 ──
    vk = L("vkospi")
    fg = L("fear_greed")
    lines = []
    if vk.get("value") is not None:
        lines.append(f"  VKOSPI {vk['value']} ({pct(vk.get('change_pct'))}) {vk.get('status', '')}")
    us = fg.get("us") or {}
    cr = fg.get("crypto") or {}
    if us.get("score") is not None:
        lines.append(f"  美 공포탐욕 {us['score']:.0f} ({us.get('rating', '')}) · 크립토 {cr.get('value', '—')} ({cr.get('rating', '')})")
    if lines:
        S.append("😱 시장 심리\n" + "\n".join(lines))

    # ── 추세추종 신호 (9종목) ──
    ts = L("trend_signal")
    lines = []
    emo = {"green": "🟢", "yellow": "🟡", "red": "🔴"}
    for t in (ts.get("tickers") or []):
        lines.append(f"  {emo.get(t.get('verdict_color'), '⚪')} {t.get('name')}: {t.get('verdict')}")
    if lines:
        S.append("🧭 추세추종 신호\n" + "\n".join(lines))

    # ── 공매도 잔고 특이 (주의/경고 + 5일 급증만) ──
    si = L("short_interest")
    lines = []
    for s in (si.get("stocks") or []):
        if s.get("signal") in ("caution", "warning"):
            tr = "▲" if s.get("trend") == "up" else "▼" if s.get("trend") == "down" else "—"
            lines.append(f"  {'🔴' if s['signal'] == 'warning' else '🟡'} {s['name']} {s.get('ratio')}% ({tr}{s.get('chg_5d'):+}%p/5일)")
        elif s.get("trend") == "up" and (s.get("chg_5d") or 0) >= 0.5:
            lines.append(f"  ⚠️ {s['name']} 잔고 급증 {s.get('chg_5d'):+}%p/5일 (현재 {s.get('ratio')}%)")
    if lines:
        S.append(f"🩳 공매도 특이 (T+2 · {si.get('asof', '')})\n" + "\n".join(lines))

    # ── 개인 수급 (감시 전 종목) ──
    # 이전엔 70% 이상만 보여줘서 보유·관심 종목이 통째로 빠졌다(SK하이닉스 45%·로보티즈 30% 등).
    # '물림이 적다'는 것도 매물대 부담이 가볍다는 유의미한 정보이므로 전 종목을 신호등으로 표시한다.
    pb = L("personal_buy_dist")
    rows = sorted((s for s in (pb.get("stocks") or []) if s.get("above_frac") is not None),
                  key=lambda s: -s["above_frac"])
    if rows:
        def _pl(s):
            f = s["above_frac"]
            ic = "🔴" if f >= 70 else "🟡" if f >= 40 else "🟢"
            va = s.get("vs_avg_pct")
            vs = f" · 평단대비 {va:+}%" if va is not None else ""
            return f"  {ic} {s['name']} 물린물량 {f}%{vs}"
        heavy = sum(1 for s in rows if s["above_frac"] >= 70)
        S.append(f"👥 개인 물림 현황 (매물대 부담 · {heavy}종목 주의)\n" + "\n".join(_pl(s) for s in rows)
                 + "\n  ※ 물린물량=현재가 위에서 개인이 순매수한 비율 — 높을수록 반등 시 매물 압박")

    # ── 밸류에이션 ──
    kv = L("kospi_valuation")
    if kv.get("per") is not None:
        _r2 = lambda v: (f"{float(v):.2f}".rstrip("0").rstrip(".") if v is not None else "—")
        S.append(f"📐 코스피 밸류 ({kv.get('asof', '')})\n"
                 f"  후행PER {_r2(kv.get('per'))} · PBR {_r2(kv.get('pbr'))} · 배당 {_r2(kv.get('div_yield'))}%")

    # ── 예측 성적 자기평가 ──
    # 제목만 싣고 본문을 비워 두면 "PQI 58.4"라는 숫자 하나만 남아 아무 판단도 못 한다.
    # prediction_quality.py 가 이미 진단·권고까지 만들어 두므로 그것을 함께 보인다.
    pq = L("prediction_quality")
    if pq.get("pqi_overall") is not None:
        head = (f"🎯 예측 자기평가: PQI {pq['pqi_overall']} {pq.get('grade', '')}"
                + (f" (추세 {pq.get('trend'):+})" if pq.get("trend") is not None else ""))
        lines = []
        dg = pq.get("diagnosis") or {}
        if dg:
            lines.append(f"  방향 {dg.get('dir_avg', '—')}% · 종가오차 {dg.get('mae_avg', '—')}% · "
                         f"밴드적중 {dg.get('range_avg', '—')}% · 급오차일 {dg.get('big_miss_avg', '—')}%")
        # 모델별 성적 — 어느 모델이 발목을 잡는지 한 줄로
        ms = [m for m in (pq.get("models") or []) if isinstance(m, dict) and m.get("pqi") is not None]
        if ms:
            ms.sort(key=lambda m: m["pqi"], reverse=True)
            lines.append("  " + " · ".join(f"{m.get('name')} {m['pqi']}" for m in ms[:4]))
        # 자동진단 권고 — 우선순위 순, 자동적용 가능 여부까지
        for r in (pq.get("recommendations") or [])[:2]:
            mark = "🔧" if r.get("auto_safe") else "⚠️"
            lines.append(f"  {mark} {r.get('issue')} → {r.get('fix')}")
        dv = pq.get("delta_vs_prev_version")
        if dv is not None:
            lines.append(f"  이전 모델 버전 대비 {dv:+}p")
        S.append(head + ("\n" + "\n".join(lines) if lines else ""))

    # ── 오늘·내일 일정 ──
    ec = L("economic_calendar")
    today_s = now.strftime("%Y-%m-%d")
    tomorrow_s = (now + timedelta(days=1)).strftime("%Y-%m-%d")
    evs = [e for e in (ec.get("events") or []) if e.get("date") in (today_s, tomorrow_s)]
    evs.sort(key=lambda e: (e.get("date"), 0 if e.get("impact") == "HIGH" else 1))
    lines = [f"  {'🔥' if e.get('impact') == 'HIGH' else '·'} [{('오늘' if e.get('date') == today_s else '내일')}] {e.get('title')}" for e in evs[:6]]
    if lines:
        S.append("📅 오늘·내일 주요 일정\n" + "\n".join(lines))

    # ── 이번 달 남은 주요 일정 (HIGH 임팩트 중심 — 미리 대비용) ──
    mon = now.strftime("%Y-%m")
    ahead = [e for e in (ec.get("events") or [])
             if str(e.get("date", "")).startswith(mon) and e.get("date") > tomorrow_s]
    ahead.sort(key=lambda e: e.get("date"))
    hi = [e for e in ahead if e.get("impact") == "HIGH"][:7]
    rest = [e for e in ahead if e.get("impact") != "HIGH"][:3]
    if hi or rest:
        def _fmt(e):
            d = str(e.get("date", ""))
            dd = f"{int(d[5:7])}/{int(d[8:10])}"
            return f"  {'🔥' if e.get('impact') == 'HIGH' else '·'} {dd} {e.get('title')}"
        body = "\n".join(_fmt(e) for e in hi + rest)
        S.append(f"🗓️ {now.month}월 남은 주요 일정 ({len(ahead)}건 중 상위)\n" + body)

    # ── 파생 만기일 (규칙 기반 자동 계산) ──
    try:
        from core.expiry import upcoming as _upcoming
        exp = _upcoming(today=now.date(), months=4)[:4]
    except Exception as e:
        print(f"  [WARN] 만기일 계산 실패: {e}")
        exp = []
    if exp:
        WDK = "월화수목금토일"

        def _fx(x):
            y, m, dd = (int(v) for v in x["date"].split("-"))
            wd = WDK[date(y, m, dd).weekday()]
            mark = "🔥" if (x["quad"] or x["d_day"] <= 3) else "·"
            near = "  ← 임박" if x["d_day"] <= 3 else ""
            adj = " (휴일보정)" if x.get("holiday_adjusted") else ""
            return f"  {mark} D-{x['d_day']:<3} {m}/{dd}({wd}) {x['region']} {x['title']}{adj}{near}"
        lines = [_fx(x) for x in exp]
        nxt = exp[0]
        if nxt["d_day"] <= 5:                       # 임박 시 대응 포인트 한 줄 부연
            lines.append(f"  ↳ {nxt['note']}")
        lines.append("  ※ 만기일=韓 둘째 목요일 · 美 셋째 금요일. 휴장일(설·추석 등 음력 공휴일, "
                     "美 Good Friday 포함)이면 직전 거래일로 자동 보정")
        S.append("⏳ 파생 만기일\n" + "\n".join(lines))

    # ── 주요 뉴스 톱3 ──
    mn = L("market_news")
    items = (mn.get("items") or [])[:3]
    if items:
        S.append("📰 주요 뉴스\n" + "\n".join(f"  • {(it.get('title') or '')[:70]}" for it in items))

    # ── 데이터 신선도 경고 ──
    fr = L("freshness")
    if fr.get("stale_count"):
        S.append(f"🚨 데이터 정체 {fr['stale_count']}건 — 대시보드 배너 확인")

    S.append(f"🔗 대시보드: {SITE}\n※ 자동 생성 요약 · 투자자문 아님")

    report = "\n\n".join(S)
    print(report)
    print(f"\n[길이] {len(report)}자 · 섹션 {len(S)}개")

    # ── 대시보드용 JSON 저장 (섹션 구조 + 최근 7일 아카이브) ──
    try:
        sections = []
        for s in S:
            p = s.split("\n")
            sections.append({"title": p[0], "lines": [x for x in p[1:] if x.strip()]})
        today_str = now.strftime("%Y-%m-%d")
        out_path = os.path.join(DOCS, "morning_report.json")
        archive = []
        try:                                              # 기존 리포트를 아카이브로 밀어넣기
            with open(out_path, encoding="utf-8") as f:
                prev = json.load(f)
            if prev.get("date") and prev["date"] != today_str:
                archive.append({"date": prev["date"], "weekday": prev.get("weekday", ""),
                                "sections": prev.get("sections", [])})
            archive += [a for a in (prev.get("archive") or []) if a.get("date") != today_str]
        except Exception:
            pass
        payload = {
            "generated_at": now.strftime("%Y-%m-%d %H:%M:%S KST"),
            "date": today_str, "weekday": WD[now.weekday()],
            "sections": sections, "archive": archive[:7],
            "note": "매일 아침(KST 06~08시) 자동 생성 — 대시보드 수집 데이터의 규칙기반 종합 요약(LLM 미사용). 텔레그램과 동일 내용. 투자자문 아님.",
        }
        os.makedirs(DOCS, exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
        print(f"[OK] {out_path} (섹션 {len(sections)} · 아카이브 {len(archive[:7])}일)")
    except Exception as e:
        print(f"[WARN] JSON 저장 실패: {e}")

    # 텔레그램 발송 (4096자 제한 → 섹션 경계로 분할)
    try:
        from core import send_message, get_secret
        if not get_secret("TELEGRAM_FINANCE_BOT_TOKEN"):
            print("[INFO] 텔레그램 토큰 없음 — 발송 생략(보고서 출력만).")
            return 0
        chunks, cur = [], ""
        for sec in S:
            if len(cur) + len(sec) + 2 > 3800:
                chunks.append(cur)
                cur = sec
            else:
                cur = (cur + "\n\n" + sec) if cur else sec
        if cur:
            chunks.append(cur)
        ok = all(send_message(c) for c in chunks)
        print(f"[{'OK' if ok else 'WARN'}] 텔레그램 {len(chunks)}건 발송{'' if ok else ' (일부 실패)'}")
    except Exception as e:
        print(f"[WARN] 발송 실패: {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
