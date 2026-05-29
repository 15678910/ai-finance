# 위협 모델 (THREAT_MODEL)

**목적**: ai-finance(자동화 시장 분석·모니터링) 프로젝트의 자산(asset)·위협(threat)·완화(mitigation)를 명문화. CLAUDE.md PART 10 자산.

**갱신 원칙**: 도메인 변경·새 위협 발견 시 즉시 업데이트. `tests/test_threat_scenarios.py`와 1:1 대응.

---

## 1. 시스템 개요

- **도메인**: GitHub Actions 기반 24/7 자동 시장 분석, AI 포트폴리오 분석, Telegram 실시간 알림 (시뮬레이션/분석 전용 — 자동 매매 절대 금지)
- **기술 스택**: Python, GitHub Actions (CI/CD), GitHub Pages (대시보드), Telegram Bot API, DART API, defusedxml
- **배포 환경**: GitHub Actions (자동화), GitHub Pages (정적 대시보드)
- **사용자 클래스**: anonymous(대시보드 조회자) / operator(GitHub Secrets 접근 가능한 운영자)

---

## 2. 자산 (Asset Inventory)

| ID | 자산 | 민감도 | 우선순위 | 위치 |
|----|----|------|---------|----|
| A1 | TELEGRAM_BOT_TOKEN | Critical | C | GitHub Secrets |
| A2 | DART_API_KEY | High | C | GitHub Secrets |
| A3 | Claude API Key (AI 분석용) | High | C | GitHub Secrets |
| A4 | 분석 상태 파일 (docs/*.json, *_state.json) | Medium | I | GitHub Pages / docs/ |
| A5 | GitHub Pages 대시보드 (dashboard.html) | Low | A | GitHub Pages |
| A6 | 포트폴리오 분석 결과 (투자 시그널) | Medium | I | docs/ JSON 파일 |

---

## 3. 신뢰 경계 (Trust Boundaries)

| 경계 | 외부 (untrusted) | 내부 (trusted) | 검증 메커니즘 |
|------|---------------|-------------|------------|
| B1 | 익명 대시보드 조회자 | GitHub Actions runner | GitHub Secrets (runner 전용) |
| B2 | GitHub Actions | Telegram Bot API | TELEGRAM_BOT_TOKEN 환경변수 |
| B3 | GitHub Actions | DART API | DART_API_KEY 환경변수 |
| B4 | GitHub Actions | Claude API | Claude API Key 환경변수 |

---

## 4. 위협 카탈로그 (Threat Catalog)

### A. 인증·세션 (Auth & Session)

| ID | 위협 | 자산 | 가능성 | 영향 | 완화 | 회귀 테스트 |
|----|----|----|------|----|----|----------|
| T-A1 | GitHub Actions 권한 과다 부여 (GITHUB_TOKEN 쓰기 권한) | A4, A5 | Med | High | workflow permissions 최소 권한 설정 (read/write 분리) | test_a1_github_token_scope |
| T-A2 | GitHub Actions workflow 외부 PR injection | A1~A3 | Low | Critical | pull_request_target 사용 시 secrets 접근 차단 | test_a2_workflow_injection |

### B. 권한·Scope

| ID | 위협 | 자산 | 가능성 | 영향 | 완화 | 회귀 테스트 |
|----|----|----|------|----|----|----------|
| T-B1 | 대시보드 JSON 직접 조작 (GitHub Pages 정적 파일 위조) | A4, A6 | Low | Med | GitHub Pages: GitHub Actions runner만 push 가능 (branch protection) | test_b1_dashboard_json_tampering |
| T-B2 | Telegram Bot 외부 메시지 스푸핑 | A1 | Low | Med | Bot chat_id 화이트리스트 검증 | test_b2_telegram_spoof |

### C. Secret·Credential

| ID | 위협 | 자산 | 가능성 | 영향 | 완화 | 회귀 테스트 |
|----|----|----|------|----|----|----------|
| T-C1 | TELEGRAM_BOT_TOKEN / DART_API_KEY GitHub Secrets 누출 | A1, A2 | Med | Critical | GitHub Secrets 전용, 로그 마스킹, 코드 하드코딩 금지 | test_c1_no_secrets_in_logs |
| T-C2 | 분석 스크립트에 하드코딩된 API 키 | A2, A3 | Low | Critical | git grep 정기 검사 + lint | test_c2_no_hardcoded_api_keys |
| T-C3 | GitHub Actions 로그에 시크릿 출력 | A1~A3 | Med | High | `::add-mask::` 적용 또는 로그 출력 시 변수 직접 참조 금지 | test_c3_secret_log_masking |

### D. Input·Injection

| ID | 위협 | 자산 | 가능성 | 영향 | 완화 | 회귀 테스트 |
|----|----|----|------|----|----|----------|
| T-D1 | XML 외부 엔티티(XXE) — DART XML 파싱 | A2, A4 | Low (이미 완화) | High | defusedxml 라이브러리 사용 (이미 적용) | test_d1_xxe_dart_xml |
| T-D2 | Claude API prompt injection (분석 보고서 생성) | A3, A6 | Med | Med | 시스템 프롬프트 격리, 외부 데이터 샌드박싱 | test_d2_prompt_injection |
| T-D3 | RSS 피드 악성 콘텐츠 파싱 | A4 | Low | Med | feedparser 안전 파싱, HTML 이스케이프 | test_d3_rss_malicious_content |

### E. Data Integrity

| ID | 위협 | 자산 | 가능성 | 영향 | 완화 | 회귀 테스트 |
|----|----|----|------|----|----|----------|
| T-E1 | 분석 상태 파일 사후 조작 (투자 시그널 위조) | A4, A6 | Low | High | GitHub 커밋 히스토리 + branch protection | test_e1_state_file_tampering |
| T-E2 | HMM 레짐 감지 모델 입력 데이터 오염 | A6 | Low | Med | 입력 데이터 범위 검증 | test_e2_hmm_input_validation |

### F. Rate / Abuse

| ID | 위협 | 자산 | 가능성 | 영향 | 완화 | 회귀 테스트 |
|----|----|----|------|----|----|----------|
| T-F1 | GitHub Actions 분당 과다 실행 (cron 오설정) | A1~A3 | Low | High | cron 스케줄 검토, Actions 실행 횟수 모니터링 | test_f1_actions_overcall |
| T-F2 | Telegram Bot 스팸 알림 (무한 알림 루프) | A1 | Med | Med | 중복 알림 방지 로직 (*_state.json deduplication) | test_f2_telegram_spam |

### G. Domain-Specific (금융 분석 무결성)

| ID | 위협 | 자산 | 가능성 | 영향 | 완화 | 회귀 테스트 |
|----|----|----|------|----|----|----------|
| T-G1 | AI 분석 결과 오류 → 잘못된 매수 시그널 Telegram 발송 | A1, A6 | High | High | 시뮬레이션 전용 면책 문구 필수, 자동 매매 절대 금지 (README) | test_g1_false_buy_signal_disclaimer |
| T-G2 | 가치 스크리너 결과 의도적 편향 (종목 조작) | A6 | Low | High | 알고리즘 파라미터 버전 관리 + 검증 | test_g2_screener_bias |
| T-G3 | 분석 결과를 실제 투자 결정에 사용 (설계 오용) | A6 | Med | Med | README 주의사항 명시, 분석 결과에 시뮬레이션 워터마크 | test_g3_investment_misuse_prevention |

---

## 5. 완화 매트릭스 (Mitigation Matrix)

| 위협 | Layer 1 (예방) | Layer 2 (탐지) | Layer 3 (복구) |
|------|------------|------------|------------|
| T-C1 | GitHub Secrets 전용 | Actions 로그 감사 | 토큰 rotate |
| T-D1 | defusedxml 사용 | XML 파싱 에러 로그 | 피드 소스 차단 |
| T-F2 | state 파일 deduplication | 알림 횟수 모니터링 | Bot 임시 비활성화 |
| T-G1 | 면책 문구 + 자동 매매 차단 | 시그널 이상 탐지 | 분석 결과 취소 알림 |
| T-A2 | PR 권한 분리 | workflow 실행 감사 | Secrets rotate |

---

## 6. 미결 위협 (Open / DEFER)

| ID | 위협 | 사유 | 예상 시간 | 우선순위 |
|----|----|----|---------|------|
| T-A1 | GitHub Actions 최소 권한 감사 | 다수 workflow 점검 필요 | 2h | High |
| T-B2 | Telegram chat_id 화이트리스트 | 미구현 | 1h | Med |
| T-C3 | 로그 마스킹 (`::add-mask::`) 전수 적용 | 다수 스크립트 점검 | 3h | High |
| T-E1 | branch protection rule 설정 | GitHub 설정 레벨 | 0.5h | Med |

---

## 7. 위협-자산 매트릭스 (heat map)

| 자산 \ 위협 | T-A2 | T-C1 | T-C3 | T-D1 | T-E1 | T-G1 | T-G3 |
|----------|------|------|------|------|------|------|------|
| A1 (Bot Token) | 🔴 | 🔴 | 🔴 | - | - | 🟠 | - |
| A2 (DART Key) | 🔴 | 🔴 | 🟠 | 🟠 | - | - | - |
| A3 (Claude Key) | 🔴 | 🟠 | 🟠 | 🟠 | - | - | - |
| A4 (상태 파일) | - | - | - | 🟠 | 🟠 | - | - |
| A5 (대시보드) | - | - | - | - | 🟠 | - | - |
| A6 (투자 시그널) | - | - | - | - | 🔴 | 🔴 | 🟠 |

범례: 🔴 Critical/High · 🟠 Medium · 🟡 Low

---

## 8. 변경 이력

| 날짜 | 변경 | 작성 |
|------|----|----|
| 2026-05-30 | 초기 작성 — Glasswing 영감 P1-D, defusedxml XXE 완화 반영 | P1-D |
