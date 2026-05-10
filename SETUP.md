# SETUP — API 키 발급 및 GitHub Secrets 등록

이 프로젝트를 운영하려면 3개의 무료 API 키가 필요합니다.

## 1. Telegram Bot 토큰 (필수)

긴급 알림과 일일 분석 결과를 받기 위해 필요합니다.

### 봇 생성

1. Telegram에서 [@BotFather](https://t.me/botfather) 검색
2. `/newbot` 명령어 입력
3. 봇 이름 + username 설정
4. 받은 토큰 복사 (예: `1234567890:ABC...`)

### Chat ID 확인

1. 본인 봇과 대화 시작 → 아무 메시지 전송
2. 브라우저에서 접속:
   ```
   https://api.telegram.org/bot{YOUR_TOKEN}/getUpdates
   ```
3. JSON 응답에서 `"chat":{"id": 12345678, ...}` 확인
4. 그 숫자가 Chat ID

## 2. DART API 키 (한국 공시 데이터)

분기 실적 분석에 필요합니다.

1. [OPEN DART](https://opendart.fss.or.kr/) 접속
2. **인증키 신청** 클릭 (회원가입)
3. 이메일 + 비밀번호 + API 사용 정보 입력
4. 즉시 발급 (40자 영숫자)

## 3. GitHub Secrets 등록

GitHub Actions에서 사용하기 위해 저장소 Secrets에 등록.

### 위치
```
https://github.com/15678910/ai-finance/settings/secrets/actions
```

### 등록할 3개 Secret

| Name | Value |
|---|---|
| `TELEGRAM_FINANCE_BOT_TOKEN` | (1번에서 받은 봇 토큰) |
| `TELEGRAM_FINANCE_CHAT_ID` | (1번에서 확인한 chat_id) |
| `DART_API_KEY` | (2번에서 받은 인증키) |

### 등록 절차

1. **"New repository secret"** 클릭
2. Name 입력 (정확히 위 표대로)
3. Secret 값 붙여넣기
4. **"Add secret"** 클릭
5. 3개 모두 반복

## 4. 로컬 테스트 (선택)

로컬에서 직접 실행하려면 `config/.env` 파일을 생성:

```bash
mkdir -p config
cat > config/.env << 'EOF'
TELEGRAM_FINANCE_BOT_TOKEN=your_token_here
TELEGRAM_FINANCE_CHAT_ID=12345678
DART_API_KEY=your_dart_key_here
EOF
```

> ⚠️ **`config/.env`는 절대 git에 커밋하지 마세요.** `.gitignore`에 이미 등록되어 있으나, 한 번 더 확인 권장:
> ```bash
> git check-ignore -v config/.env
> ```

## 5. 검증

GitHub Actions 탭에서 워크플로 수동 실행:

1. `Daily Finance Analysis` 선택
2. **"Run workflow"** → master 브랜치 → 실행
3. 약 5분 후 텔레그램 알림 도착 확인
4. https://15678910.github.io/ai-finance/ 에서 대시보드 갱신 확인

## 문제 해결

### 텔레그램 알림이 안 옴
- Bot Token이 정확한지 확인
- 본인이 봇과 한 번이라도 대화했는지 (Chat ID 발급 조건)
- GitHub Secrets 이름 정확히 일치하는지

### DART API 401 에러
- 키가 만료되었을 가능성 → OPEN DART에서 재발급
- 키에 공백/특수문자 섞였는지 확인

### Telegram Forbidden 403 에러
- 봇이 차단당한 상태 → 봇 메뉴 → 차단 해제

## 참고

- 모든 API는 **무료** (개인 사용량 한도 내)
- GitHub Actions는 **공개 저장소 무제한**
- yfinance/CoinGecko/FRED는 인증 불필요
