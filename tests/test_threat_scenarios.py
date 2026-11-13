"""Threat scenario regression tests — ai-finance (AI 금융 분석 시스템).

각 위협이 실제 코드 변경으로 재발하지 않도록 회귀 차단.
새 위협 발견 → 이 파일에 시나리오 추가 → fail → fix → pass.

CLAUDE.md PART 10 정책에 따른 자산.

카테고리:
- A: 인증·세션 (Telegram bot 인증, API 키 보호)
- B: 권한·scope (접근 제어, allowlist)
- C: secret·credential (env 누락, Telegram token, API key hardcoding)
- D: input·injection (메시지 파싱, 파라미터 조작)
- E: data integrity (금융 데이터 무결성, 이중 알림)
- F: rate/abuse (Telegram 메시지 스팸, 외부 API 남용)
- G: domain-specific (금융 분석 위협: 잘못된 데이터로 투자 결정, secret 누출)

실행: pytest tests/test_threat_scenarios.py -v
"""
import os
import re
import sys
import pytest
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
SKIP_DIRS = {"node_modules", ".git", "dist", "build", "venv", ".venv",
             "__pycache__", "backups", "output", "qa_output", "files.zip"}


# ============================================================
# A. 인증·세션 (Authentication & Session)
# ============================================================

class TestThreatA_Authentication:
    """인증 우회·약화 위협의 회귀 차단."""

    @pytest.mark.skip(reason="TODO: Telegram bot allowlist 미적용 시 외부인 명령 실행 가능 여부 확인")
    def test_a1_telegram_bot_allowlist_enforced(self):
        """Telegram bot이 allowlist에 없는 사용자의 명령을 거부.

        ai-finance 핵심: 봇 명령이 외부 공격자에게 노출되면 금융 정보 유출.
        """
        pass

    @pytest.mark.skip(reason="TODO: TELEGRAM_TOKEN 미설정 시 fail-fast 확인")
    def test_a2_telegram_token_missing_raises(self):
        """TELEGRAM_TOKEN 환경변수 미설정 시 서비스가 시작 전 예외 발생.

        core.env_loader의 get_secret() 동작 확인.
        """
        pass


# ============================================================
# B. 권한·Scope (Authorization & Scope)
# ============================================================

class TestThreatB_Authorization:
    """권한 상승·데이터 접근 차단."""

    @pytest.mark.skip(reason="TODO: Telegram chat_id allowlist 검증 로직 확인")
    def test_b1_unauthorized_chat_id_blocked(self):
        """허용되지 않은 Telegram chat_id로 명령 전송 시 무시 또는 거부."""
        pass


# ============================================================
# C. Secret·Credential
# ============================================================

class TestThreatC_Secrets:
    """hardcoded secret·env 누출·평문 비교 차단."""

    def test_c1_no_hardcoded_api_keys_in_source(self):
        """32자 hex 패턴(Telegram token 일부, 증권사 API key 등)이 소스에 없어야 함.

        ai-finance: Telegram bot token, 증권사 API key, Yahoo Finance 키 노출 위험.
        """
        hex32 = re.compile(r'[0-9a-f]{32}', re.IGNORECASE)
        found = []
        for root, dirs, files in os.walk(REPO_ROOT):
            dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
            for f in files:
                if not f.endswith(('.py', '.js', '.jsx', '.ts', '.tsx')):
                    continue
                path = os.path.join(root, f)
                try:
                    content = Path(path).read_text(encoding='utf-8', errors='ignore')
                    for m in hex32.finditer(content):
                        idx = m.start()
                        ctx = content[max(0, idx - 50):idx + 82].lower()
                        if any(x in ctx for x in ('todo', 'example', 'placeholder',
                                                   'your-', 'test', 'dummy', 'fake',
                                                   'sha256', 'hash', 'md5', 'hex',
                                                   'ticker', 'symbol')):
                            continue
                        found.append((path.replace(str(REPO_ROOT), ''), m.group()[:8] + '...'))
                except Exception:
                    pass
        assert not found, f"Possible hardcoded API/secret keys found: {found}"

    def test_c2_no_telegram_bot_token_pattern_in_source(self):
        """Telegram bot token 패턴 (숫자:alphanumeric 35자)이 소스에 없어야 함."""
        # Telegram token format: 123456789:ABCdefGHIjklMNOpqrSTUvwxyz1234567
        tg_token = re.compile(r'\d{8,10}:[A-Za-z0-9_-]{35}')
        found = []
        for root, dirs, files in os.walk(REPO_ROOT):
            dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
            for f in files:
                if not f.endswith(('.py', '.js', '.jsx', '.ts', '.tsx', '.env')):
                    continue
                if f.endswith('.env'):
                    continue  # .env 파일은 gitignore 대상, 여기선 제외
                path = os.path.join(root, f)
                try:
                    content = Path(path).read_text(encoding='utf-8', errors='ignore')
                    for m in tg_token.finditer(content):
                        ctx = content[max(0, m.start() - 30):m.start() + 70].lower()
                        if any(x in ctx for x in ('todo', 'example', 'placeholder', 'your_bot')):
                            continue
                        found.append((path.replace(str(REPO_ROOT), ''), m.group()[:12] + '...'))
                except Exception:
                    pass
        assert not found, f"Telegram bot token pattern found in source: {found}"

    def test_c3_required_env_vars_documented(self):
        """필수 env var(TELEGRAM_TOKEN, 증권사 API key 등)가 문서화되어 있는지.

        위협: 운영자가 필요한 secret을 모른 채 배포 → 해당 모듈이 조용히
        스킵되거나 빈 데이터를 생성하고, 그 결과로 잘못된 투자 판단.
        소스가 실제로 읽는 secret은 모두 SETUP.md/README.md에 기재되어야 함.
        """
        # get_secret("X") / os.environ.get("X") / os.environ["X"] / os.getenv("X")
        patterns = [
            re.compile(r'get_secret\(\s*["\']([A-Z][A-Z0-9_]{2,})["\']'),
            re.compile(r'os\.environ\.get\(\s*["\']([A-Z][A-Z0-9_]{2,})["\']'),
            re.compile(r'os\.environ\[\s*["\']([A-Z][A-Z0-9_]{2,})["\']'),
            re.compile(r'os\.getenv\(\s*["\']([A-Z][A-Z0-9_]{2,})["\']'),
        ]
        # secret이 아닌 런타임 환경변수는 문서화 대상에서 제외
        RUNTIME_VARS = {
            "PATH", "HOME", "TZ", "LANG", "LC_ALL", "TEMP", "TMP", "TMPDIR",
            "PYTHONPATH", "PYTHONIOENCODING", "PYTHONUTF8", "VIRTUAL_ENV",
            "CI", "GITHUB_ACTIONS", "GITHUB_WORKSPACE", "GITHUB_REPOSITORY",
            "GITHUB_REF", "GITHUB_SHA", "GITHUB_RUN_ID", "GITHUB_OUTPUT",
            "GITHUB_ENV", "RUNNER_OS", "USERPROFILE", "APPDATA",
        }

        used = {}  # var -> 처음 발견된 파일
        for root, dirs, files in os.walk(REPO_ROOT):
            dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
            if os.path.basename(root) == "tests":
                continue  # 테스트 자신의 참조는 제외
            for f in files:
                if not f.endswith(".py"):
                    continue
                path = os.path.join(root, f)
                try:
                    content = Path(path).read_text(encoding="utf-8", errors="ignore")
                except OSError:
                    continue
                for pat in patterns:
                    for m in pat.finditer(content):
                        name = m.group(1)
                        if name in RUNTIME_VARS:
                            continue
                        used.setdefault(name, path.replace(str(REPO_ROOT), ""))

        assert used, "env var 스캔 결과가 비어 있음 — 스캔 로직 자체를 점검할 것"

        # 문서 소스: SETUP.md, README.md, *.env.example
        doc_text = ""
        doc_names = []
        for candidate in ("SETUP.md", "README.md", ".env.example",
                          "config/.env.example"):
            p = REPO_ROOT / candidate
            if p.exists():
                doc_text += p.read_text(encoding="utf-8", errors="ignore")
                doc_names.append(candidate)
        assert doc_names, "문서 파일(SETUP.md/README.md)이 없음"

        missing = sorted((v, src) for v, src in used.items() if v not in doc_text)
        assert not missing, (
            f"소스에서 사용하지만 {doc_names}에 문서화되지 않은 env var: {missing}"
        )


# ============================================================
# D. Input·Injection
# ============================================================

class TestThreatD_Injection:
    """메시지 파싱 오류, 파라미터 조작 차단."""

    def test_d1_telegram_message_not_evaluated(self):
        """Telegram 메시지 내용이 eval/exec로 처리되지 않음.

        1) 소스 전체에 코드 실행 sink가 없어야 함 (메시지·외부 데이터가
           흘러들 경로를 원천 차단).
        2) send_message가 본문을 URL이 아닌 POST body에 urlencode해 넣어,
           메시지에 섞인 `&chat_id=...`가 API 파라미터로 승격되지 않아야 함.
        """
        # ── 1) 정적 스캔: 코드 실행 sink ──────────────────────────────
        # __import__는 제외: 하드코딩된 의존성 목록 점검(check_and_install)에
        # 쓰이며 외부 입력이 닿지 않음.
        dangerous = [
            (re.compile(r'(?<![\w.])eval\s*\('), "eval()"),
            (re.compile(r'(?<![\w.])exec\s*\('), "exec()"),
            (re.compile(r'os\.system\s*\('), "os.system()"),
            (re.compile(r'shell\s*=\s*True'), "shell=True"),
            (re.compile(r'pickle\.loads?\s*\('), "pickle.load(s)"),
            (re.compile(r'new\s+Function\s*\('), "new Function()"),
            (re.compile(r'setTimeout\s*\(\s*["\']'), "setTimeout(문자열)"),
        ]
        yaml_load = re.compile(r'yaml\.load\s*\(')
        safe_loader = re.compile(r'Safe(Loader|Load)')

        found = []
        scanned = 0
        for root, dirs, files in os.walk(REPO_ROOT):
            dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
            if os.path.basename(root) == "tests":
                continue  # 테스트 자신의 패턴 문자열은 제외
            for f in files:
                if not f.endswith((".py", ".js", ".html")):
                    continue
                path = os.path.join(root, f)
                try:
                    content = Path(path).read_text(encoding="utf-8", errors="ignore")
                except OSError:
                    continue
                scanned += 1
                rel = path.replace(str(REPO_ROOT), "")
                for lineno, line in enumerate(content.splitlines(), 1):
                    for pat, label in dangerous:
                        if pat.search(line):
                            found.append(f"{rel}:{lineno} {label}")
                    if yaml_load.search(line) and not safe_loader.search(line):
                        found.append(f"{rel}:{lineno} yaml.load(비-Safe)")

        assert scanned > 0, "스캔된 파일이 없음 — 스캔 로직 자체를 점검할 것"
        assert not found, f"코드 실행 sink 발견 (외부 입력 유입 시 RCE): {found}"

        # ── 2) 기능: 메시지가 API 파라미터로 승격되지 않는지 ───────────
        import urllib.parse
        import urllib.request

        from core import send_message

        captured = {}

        class _FakeResponse:
            def read(self):
                return b'{"ok": true}'

            def __enter__(self):
                return self

            def __exit__(self, *exc):
                return False

        def _fake_urlopen(req, *args, **kwargs):
            captured["url"] = req.full_url
            captured["body"] = req.data
            return _FakeResponse()

        # 메시지에 API 파라미터 주입 + 셸/템플릿 페이로드를 섞어 전송
        payload = ("/start & chat_id=999999999 & text=hacked\n"
                   "$(rm -rf /) `id` ${HOME}\n"
                   "{{7*7}} <script>alert(1)</script>\n"
                   "../../etc/passwd %0d%0aInjected: 1")

        orig = urllib.request.urlopen
        urllib.request.urlopen = _fake_urlopen
        try:
            ok = send_message(payload, bot_token="1111111111:TESTTOKEN",
                              chat_id="42")
        finally:
            urllib.request.urlopen = orig

        assert ok is True, "테스트 더블 전송 경로가 동작하지 않음"

        # URL에는 토큰/메서드만, 메시지 본문은 포함되지 않아야 함
        assert captured["url"].endswith("/sendMessage"), captured["url"]
        assert "chat_id" not in captured["url"]
        assert "hacked" not in captured["url"]

        # POST body는 urlencode되어 파라미터 경계가 깨지지 않아야 함
        parsed = urllib.parse.parse_qs(captured["body"].decode("utf-8"),
                                       keep_blank_values=True)
        assert parsed["chat_id"] == ["42"], \
            f"메시지의 chat_id가 실제 파라미터를 덮어씀: {parsed.get('chat_id')}"
        assert parsed["text"] == [payload], \
            "메시지 본문이 전송 과정에서 해석·변형됨 (주입 위험)"
        # 개행이 헤더로 승격되지 않도록 인코딩되어야 함 (CRLF injection)
        assert "\n" not in captured["body"].decode("utf-8")

    @pytest.mark.skip(reason="TODO: 티커 심볼 입력이 허용 문자로만 구성되는지 확인")
    def test_d2_ticker_symbol_validated(self):
        """티커 심볼이 알파벳/숫자/하이픈으로만 구성된 허용 목록에서 처리됨.

        BAD: f"https://finance.yahoo.com/quote/{user_ticker}"  (path traversal 가능)
        GOOD: ticker = sanitize(user_ticker) → whitelist check
        """
        pass


# ============================================================
# E. Data Integrity
# ============================================================

class TestThreatE_Integrity:
    """금융 데이터 무결성, 이중 알림 차단."""

    def test_e1_double_alert_prevented(self):
        """같은 이벤트에 대한 알림이 설정된 시간 내 중복 전송되지 않음.

        core.state_store.is_recent_alert 동작 검증.

        위협: 같은 시그널(예: 급락 경보)이 매 실행마다 재전송되면
        투자자가 새 이벤트로 오인해 잘못된 매매 판단을 내림.
        """
        import tempfile
        from datetime import datetime, timezone, timedelta

        from core import (load_state, save_state, is_recent_alert,
                          mark_alert_sent)

        THRESHOLD = 12

        # 1) 최초 알림: 기록 없음 → 전송 허용
        state = {}
        assert is_recent_alert(state, "kospi_crash", hours=THRESHOLD) is False

        # 2) 전송 직후: 같은 key 재전송 차단
        state = mark_alert_sent(state, "kospi_crash")
        assert is_recent_alert(state, "kospi_crash", hours=THRESHOLD) is True, \
            "이중 알림 차단 실패: 방금 보낸 알림이 다시 전송 가능 상태"

        # 3) key 격리: 다른 이벤트는 차단되면 안 됨 (알림 유실 방지)
        assert is_recent_alert(state, "samsung_earnings", hours=THRESHOLD) is False

        # 4) 프로세스 재시작(저장 → 로드) 후에도 차단 유지
        with tempfile.TemporaryDirectory() as tmpdir:
            assert save_state("threat_e1", state, state_dir=tmpdir) is True
            reloaded = load_state("threat_e1", state_dir=tmpdir)
            assert is_recent_alert(reloaded, "kospi_crash", hours=THRESHOLD) is True, \
                "이중 알림 차단 실패: 상태 파일 재로드 후 중복 전송 가능"

        # 5) 임계 시간 경계: 이내는 차단, 초과는 허용
        now = datetime.now(timezone.utc)
        within = {"last_alerts": {
            "kospi_crash": (now - timedelta(hours=THRESHOLD - 1)).isoformat()}}
        expired = {"last_alerts": {
            "kospi_crash": (now - timedelta(hours=THRESHOLD + 1)).isoformat()}}
        assert is_recent_alert(within, "kospi_crash", hours=THRESHOLD) is True
        assert is_recent_alert(expired, "kospi_crash", hours=THRESHOLD) is False, \
            "임계 시간 초과 알림이 영구 차단됨 (알림 유실 위험)"

        # 6) tz 없는 timestamp도 UTC로 해석 (로컬시간 오해로 인한 오차단 방지)
        naive = {"last_alerts": {
            "kospi_crash": (now - timedelta(hours=1)).replace(tzinfo=None).isoformat()}}
        assert is_recent_alert(naive, "kospi_crash", hours=THRESHOLD) is True

        # 7) 손상된 상태값 → fail-open (영구 차단으로 알림이 사라지면 안 됨)
        for broken in ("", "not-a-timestamp", None, 12345, {"a": 1}):
            assert is_recent_alert({"last_alerts": {"k": broken}},
                                   "k", hours=THRESHOLD) is False, \
                f"손상된 timestamp({broken!r})가 알림을 영구 차단함"

        # 8) state 자체가 dict가 아닌 경우에도 예외 없이 처리
        assert is_recent_alert(None, "kospi_crash", hours=THRESHOLD) is False
        assert is_recent_alert([], "kospi_crash", hours=THRESHOLD) is False

    @pytest.mark.skip(reason="TODO: 금융 수치(PER, 배당률 등) 이상값 필터링 확인")
    def test_e2_financial_metric_outlier_filtered(self):
        """비정상적으로 큰/작은 금융 수치(예: PER 10000x)가 필터링 또는 플래그됨."""
        pass


# ============================================================
# F. Rate / Abuse
# ============================================================

class TestThreatF_Rate:
    """DoS, 외부 API 남용 차단."""

    @pytest.mark.skip(reason="TODO: yfinance API 호출에 retry/backoff 로직 있는지 확인")
    def test_f1_yfinance_calls_have_retry_backoff(self):
        """Yahoo Finance API 호출 실패 시 지수 backoff retry 적용."""
        pass


# ============================================================
# G. Domain-Specific — AI 금융 분석
# ============================================================

class TestThreatG_FinanceDomain:
    """AI 금융 분석 시스템의 도메인 특화 위협.

    핵심 위협:
    - 잘못된 데이터 기반 투자 결정 (hallucination, stale data)
    - Telegram을 통한 금융 정보 외부 유출
    - 증권사 API key 노출 → 무단 거래
    - 이중 알림 → 투자자 혼란
    """

    @pytest.mark.skip(reason="TODO: 금융 데이터 freshness 확인 (stale data로 잘못된 판단 방지)")
    def test_g1_financial_data_staleness_detected(self):
        """금융 데이터의 타임스탬프가 설정된 임계값(예: 24시간)을 초과하면 경고."""
        pass

    def test_g2_investment_recommendation_has_disclaimer(self):
        """AI가 생성한 투자 추천 메시지에 '이는 투자 권유가 아닙니다' 면책 조항 포함.

        docs/THREAT_MODEL.md T-G1의 완화책("시뮬레이션 전용 면책 문구 필수,
        자동 매매 절대 금지")이 실제 산출물에 남아 있는지 회귀 차단.

        검증 대상:
          1) README.md — 면책 + 자동 매매 금지 명시
          2) docs/index.html — 공개 대시보드 면책 조항
          3) 투자의견(목표가·투자의견·비중확대 등)을 Telegram으로 보내는 모듈
        """
        disclaimer = re.compile(
            r'투자자문\s*아님|투자\s*권유가?\s*아[님닙]|투자\s*권유가\s*아닙니다|'
            r'정보\s*제공용?|정보용|참고용|not\s+investment\s+advice',
            re.IGNORECASE)

        # ── 1) README: 면책 + 자동 매매 금지 ──────────────────────────
        readme = (REPO_ROOT / "README.md").read_text(encoding="utf-8",
                                                     errors="ignore")
        assert disclaimer.search(readme), \
            "README.md에 면책 조항이 없음 (THREAT_MODEL T-G1 완화책 유실)"
        assert re.search(r'자동\s*매매\s*(절대\s*)?금지|자동\s*매매.*금지', readme), \
            "README.md에 '자동 매매 금지' 명시가 없음"

        # ── 2) 공개 대시보드 ─────────────────────────────────────────
        dashboard = REPO_ROOT / "docs" / "index.html"
        if dashboard.exists():
            html = dashboard.read_text(encoding="utf-8", errors="ignore")
            assert disclaimer.search(html), \
                "docs/index.html(공개 대시보드)에 면책 조항이 없음"

        # ── 3) 투자의견을 Telegram으로 보내는 모듈 ────────────────────
        # 수급 데이터(순매수/매도세)는 추천이 아니므로 제외하고,
        # 명시적 투자의견 표현만 대상으로 함.
        sends = re.compile(r'send_message\s*\(|send_telegram\s*\(')
        opinion = re.compile(
            r'목표가|목표\s*주가|투자의견|비중\s*(확대|축소)|추천\s*종목|'
            r'손절|진입\s*가|\bBUY\b|\bSELL\b|OVERWEIGHT|UNDERWEIGHT')

        violations = []
        for path in sorted(REPO_ROOT.glob("*.py")):
            try:
                content = path.read_text(encoding="utf-8", errors="ignore")
            except OSError:
                continue
            if not sends.search(content):
                continue
            hits = sorted({m.group() for m in opinion.finditer(content)})
            if hits and not disclaimer.search(content):
                violations.append((path.name, hits[:5]))

        assert not violations, (
            "투자의견을 Telegram으로 전송하지만 면책 문구가 없는 모듈: "
            f"{violations}"
        )

    @pytest.mark.skip(reason="TODO: 브로커 API key 없이 실제 거래 실행 불가 확인")
    def test_g3_no_auto_trade_without_explicit_broker_key(self):
        """브로커 API key 없이 자동 매매 코드 경로가 실행되지 않음."""
        pass
