# ai-finance — 프로젝트 규칙

GitHub Pages 정적 대시보드(`docs/`). 수집기(`*.py`)가 `docs/*.json`을 쓰고, `docs/index.html`이 클라이언트에서 불러온다.

## 수정 금지 영역

- `docs/*_state.json` — 수집기 중복 알림 방지 상태(`core/state_store.py`). 손으로 고치지 않는다.
- 포트폴리오 원금·보유종목·매수가 — 브라우저 localStorage에만 둔다. 공개 레포에 커밋 금지.

## 자주 발생하는 실수

| 날짜 | 실수 | 대처 |
|---|---|---|
| 2026-11-16 | 10-08 18:30 로컬 실행의 `pull --rebase`가 중단돼 `.git/rebase-merge`가 남고 HEAD가 떨어짐. 이후 5.5주 동안 detached HEAD에 커밋되어 push가 매번 실패했는데 `.bat`은 무조건 `PUSHED`를 기록함 → 공개 대시보드의 KRX 데이터 5종이 10-07에 멈춤 | `update_kospi_valuation.bat`에 중단 rebase 정리·브랜치 가드·push 결과 검사 추가. 확인: `git status -sb`가 `## master...origin/master`인지 |
| 2026-11-16 | 오래된 rebase 상태를 `git rebase --abort`로 치우면 브랜치가 rebase 이전 시점으로 돌아간다 | 이미 브랜치를 따로 맞췄다면 HEAD를 건드리지 않는 `git rebase --quit`을 쓴다 |
| 2026-11-16 | KRX 비밀번호 만료("패스워드 변경 필요") → 로컬 수집기 4개가 `KRX 로그인 거부`로 매일 중단 | 사용자가 krx.co.kr에서 변경 후 `KRX_PW` 재설정. Claude는 자격증명을 입력하지 않는다 |
| 2026-11-16 | 비밀번호를 바꾼 뒤 확인 없이 전체 실행 + 단일 수집기 시험을 돌려 실패가 약 7번 쌓임 → KRX가 아이디/비밀번호 로그인을 잠금(카카오 로그인은 정상) | 자격증명을 바꾼 뒤에는 `python krx_login_check.py --registry`로 **한 번만** 확인하고 CD001일 때만 수집기를 돌린다. `.bat`은 이제 사전 점검 1회 후 실패하면 KRX 로그인 없이 진행하고 알린다 |
| 2026-11-17 | 이전 세션이 "한국 시장 폭 지표가 없다"며 새 KRX 수집기를 제안했지만 `market_breadth.py`(시총 상위 800, GitHub Actions)가 이미 있었다 | 새 지표를 계획하기 전에 `grep -rli <주제> --include=*.py .`와 `docs/*.json`부터 확인한다 |
| 2026-11-16 | pykrx는 사용자 site-packages에 있어 `python -I`로는 안 보이고, import 시점에 이모지를 출력해 cp949 콘솔에서 죽는다 | `C:\Python313\python.exe` + `PYTHONIOENCODING=utf-8 PYTHONUTF8=1` |

## 로컬 전용 수집기

KRX·INDEXerGO는 GitHub 데이터센터 IP를 막는다. 아래는 작업 스케줄러 "KOSPI Valuation Update"(매일 18:30)가 `update_kospi_valuation.bat`으로 실행한다:
`kospi_valuation_monitor.py`, `short_interest_monitor.py`, `personal_buy_dist.py`, `stock_financials.py`, `stock_screener.py`.

## 배포 전 체크리스트

- [ ] `git status -sb`가 브랜치 위인가 (detached HEAD 아님)
- [ ] `git rev-list --left-right --count origin/master...HEAD`가 `0 0`인가 (push 후)
- [ ] 수정한 수집기를 실제로 실행해 JSON이 생성됐는가
