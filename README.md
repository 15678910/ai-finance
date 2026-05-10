# AI Finance — 자동화된 시장 분석 및 모니터링 시스템

> 🚨 **시뮬레이션 / 분석용 전용. 자동 매매 절대 금지. 실제 투자 결정은 본인의 판단 필요.**

GitHub Actions 기반 24/7 자동 운영, GitHub Pages 대시보드, Telegram 실시간 알림.

**대시보드**: https://15678910.github.io/ai-finance/

## 핵심 기능

| 영역 | 시스템 |
|---|---|
| 📊 **시장 분석** | 4시간마다 7개 섹터 22+종목 분석 (HMM 레짐, 센티멘트, 포트폴리오) |
| 🔔 **긴급 뉴스** | 1시간마다 7개 RSS 피드 + 6개 카테고리 키워드 모니터링 |
| 🌐 **해외 시장** | KOSPI 휴장 매시간 13개 글로벌 지수/선물 추적 |
| 🇯🇵 **일본 위기** | 엔화 캐리 청산 압력 / BOJ 금리 인상 추적 (매일 09:30, 22:00) |
| 🪙 **비트코인 본위제** | ETF 자금 흐름 / 채굴 / ARMA 법안 (매일 09:00, 22:00) |
| 💰 **저평가 스크리너** | KOSPI 47개 종목 가치 점수 + WATCHLIST (매일 06:00) |
| 📈 **DCF 적정가** | 15개 대형주 현금흐름 평가 (매일 07:30) |
| 📊 **분기 실적 리뷰** | DART API 기반 어닝 분석 (매일 09:00) |
| 🧬 **AutoResearch** | 가치 가중치 진화 + Forward Test 평가 (매일 + 매월) |

## 빠른 시작

### 1. 저장소 클론
```bash
git clone https://github.com/15678910/ai-finance.git
cd ai-finance
```

### 2. 의존성 설치
```bash
pip install -r requirements.txt
```

### 3. API 키 설정 (자세한 내용은 [SETUP.md](SETUP.md))
- Telegram Bot 토큰 + Chat ID
- DART API 키 (한국 공시 데이터)

### 4. 로컬 테스트
```bash
PYTHONIOENCODING=utf-8 python overseas_market_monitor.py
```

### 5. 자동 운영 활성화
GitHub에 push하면 모든 워크플로가 자동 실행됩니다.

## 프로젝트 구조

```
ai-finance/
├── core/                    # 공통 유틸리티
│   ├── env_loader.py        # .env 파싱 + 환경변수
│   ├── telegram.py          # 텔레그램 메시징
│   ├── yf_helper.py         # yfinance 래퍼 + 재시도
│   └── state_store.py       # JSON 상태 파일 관리
├── docs/                    # GitHub Pages 출력 (자동 생성)
│   ├── index.html           # 대시보드
│   ├── data.json            # 통합 데이터
│   └── *.json               # 모듈별 분석 결과
├── config/                  # 설정 (gitignored)
│   ├── .env                 # API 키 (절대 커밋 금지)
│   └── sectors.json         # 분석 섹터 정의
├── .github/workflows/       # 11개 GitHub Actions
├── tests/                   # 스모크 테스트
├── *_monitor.py             # 모니터링 스크립트
├── *_analyzer.py            # 분석 스크립트
├── value_screener.py        # 가치 스크리너
├── dcf_valuator.py          # DCF 평가기
├── earnings_reviewer.py     # 분기 실적 리뷰어
├── auto_research_*.py       # AutoResearch 시스템
└── generate_dashboard_data.py # 대시보드 데이터 빌더
```

## 자동화 시스템 (11개)

전체 운영 스케줄과 아키텍처는 [ARCHITECTURE.md](ARCHITECTURE.md) 참조.

| 워크플로 | 주기 | 핵심 출력 |
|---|---|---|
| Daily Finance Analysis | 4시간마다 | data.json (전체 통합) |
| Breaking News Monitor | 1시간마다 | 텔레그램 긴급 알림 |
| Overseas Market Monitor | KOSPI 휴장 매시간 | 해외 시장 + 시초가 예측 |
| Japan Crisis Monitor | 매일 09:30, 22:00 | 캐리 청산 압력 |
| Bitcoin Standard Monitor | 매일 09:00, 22:00 | BTS 점수 + 한국 영향 |
| Value Screener | 매일 06:00 | 저평가 Top 10 |
| Auto-Research Value | 매일 07:00 | 진화된 가중치 |
| Auto-Research Portfolio | 매일 03:00 | 포트폴리오 최적화 |
| DCF Valuator | 매일 07:30 | 적정가 vs 현재가 |
| Earnings Reviewer | 매일 09:00 | 분기 실적 점수 |
| Forward Test Eval | 매월 1일 08:00 | 실현 알파 평가 |

## 보안

- ✅ API 키는 GitHub Secrets에 저장 (`config/.env` 절대 커밋 안 함)
- ✅ `defusedxml`로 XXE 공격 방어
- ✅ 모든 외부 호출 HTTPS
- ✅ 워크플로 동시성 제어로 race condition 방지
- ✅ subprocess 호출 시 list args (shell=True 금지)

## 라이선스 / 면책

본 프로젝트는 **개인 학습 및 분석 용도**입니다. 모든 분석 결과는 시뮬레이션이며 투자 권유가 아닙니다. 실제 투자 결정은 본인의 판단으로 이루어져야 하며, 본 시스템의 결과로 인한 손실에 대해 책임지지 않습니다.

## 문서

- [SETUP.md](SETUP.md) — API 키 발급 + GitHub Secrets 등록
- [ARCHITECTURE.md](ARCHITECTURE.md) — 시스템 아키텍처 + 데이터 흐름
- [CHANGELOG.md](CHANGELOG.md) — 버전별 변경 이력
