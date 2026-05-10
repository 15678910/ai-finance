# ARCHITECTURE — 시스템 설계 및 데이터 흐름

## 전체 구조

```
┌─────────────────────────────────────────────────────────────┐
│  데이터 소스                                                  │
│  ┌─────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐    │
│  │yfinance │  │   FRED   │  │ DART API │  │CoinGecko │    │
│  │(가격)   │  │(매크로)  │  │(공시)    │  │(스테이블)│    │
│  └────┬────┘  └────┬─────┘  └────┬─────┘  └────┬─────┘    │
└───────┼────────────┼──────────────┼─────────────┼──────────┘
        │            │              │             │
        ▼            ▼              ▼             ▼
┌─────────────────────────────────────────────────────────────┐
│  GitHub Actions (11개 워크플로)                              │
│  ┌──────────────────────────────────────────────┐           │
│  │ Python 모니터/분석 스크립트                  │           │
│  │  • core/ 공통 유틸리티                       │           │
│  │  • *_monitor.py / *_analyzer.py              │           │
│  └──────────────────────────────────────────────┘           │
└─────────────────────────────────────────────────────────────┘
        │
        ▼
┌─────────────────────────────────────────────────────────────┐
│  결과물                                                      │
│  ┌──────────────────┐  ┌──────────────────┐                │
│  │ docs/*.json      │  │ Telegram 알림    │                │
│  │ (대시보드 데이터)│  │ (실시간 push)    │                │
│  └────────┬─────────┘  └──────────────────┘                │
│           │                                                  │
│           ▼                                                  │
│  ┌──────────────────────────────────┐                       │
│  │ GitHub Pages (대시보드)          │                       │
│  │ https://15678910.github.io/...   │                       │
│  └──────────────────────────────────┘                       │
└─────────────────────────────────────────────────────────────┘
```

## 핵심 모듈 (core/)

리뷰 결과 600+ LOC의 중복 코드를 발견하여 추출한 공통 유틸리티.

### `core/env_loader.py`
- `load_env(path)`: `.env` 파일 파싱 (따옴표/주석 처리)
- `get_secret(key)`: 환경변수 우선, `.env` 폴백

### `core/telegram.py`
- `send_message(text)`: 텔레그램 텍스트 (4096자 자동 절단)
- `send_document(path)`: 파일 전송 (multipart)

### `core/yf_helper.py`
- `resolve_ticker("005930")` → `"005930.KS"` (KS/KQ 자동 판별)
- `fetch_history(ticker, period)`: 재시도 + 지수 백오프
- `fetch_info_safely(ticker)`: info 안전 조회
- `get_current_price(ticker)`: history NaN 회피

### `core/state_store.py`
- `load_state(name)`: `docs/{name}_state.json` 로드
- `save_state(name, data)`: 저장
- `is_recent_alert(state, key, hours)`: 알림 중복 방지
- `mark_alert_sent(state, key)`: 발송 시각 기록

## 시스템별 역할

### 1. Daily Finance Analysis (오케스트레이터)
- **주기**: 4시간마다 (KST 00, 04, 08, 12, 16, 20)
- **역할**: 모든 분석을 통합 실행 + dashboard data 빌드
- **순서**:
  1. `daily_sector_analysis.py` (섹터 7개 × 종목별 분석)
  2. `value_screener.py` (저평가 Top 10)
  3. `auto_research_value.py` (가중치 진화)
  4. `generate_dashboard_data.py` (data.json 통합)
  5. Telegram 발송 + Git push

### 2. Breaking News Monitor
- **주기**: 1시간마다
- **데이터**: 7개 RSS 피드 (Reuters, CNBC, BBC, 연합뉴스, Google News)
- **로직**: 6개 카테고리 × 영/한 키워드 매칭 → 신규 뉴스만 알림
- **상태**: `docs/breaking_news_state.json` (seen_links 500개 유지)

### 3. Overseas Market Monitor
- **주기**: KOSPI 휴장 시간대 매시간
- **추적**: 13개 글로벌 지수 (S&P, Nasdaq, Nikkei, FTSE, DAX, VIX 등)
- **알림 트리거**: ±2% 일반, ±3% 긴급, VIX 30+

### 4. Japan Crisis Monitor
- **데이터**: yfinance + FRED (BOJ 정책금리, 일본 10Y 등)
- **계산**:
  - 캐리 압력 지수 = 엔화 1주 변동 + 일미 금리차 + 일본 10Y 추세
  - 위기 종합 점수 = 캐리 + USD/JPY 절대 + CPI + 부채/GDP

### 5. Bitcoin Standard Monitor
- **추적**: BTC, 스테이블코인 시총, BTC ETF, 채굴 종목
- **점수**: BTS = BTC강세 + 스테이블코인 채택 + 도미넌스 + AI수요
- **법안 추적**: ARMA Bill 진행 단계 + 시나리오 (낙관/중립/비관)

### 6. Value Screener
- **종목**: 47개 KOSPI/KOSDAQ 대형주
- **점수**: 밸류(35%) + 수익성(35%) + 성장(20%) + 주주환원(10%)
- **필터**: 시총 1000억+, ROE -20%+, 부채비율 500%-

### 7. AutoResearch
- **Portfolio**: 동일 가중 → ±10% 변형 → Sortino 최대화
- **Value**: 가치 가중치 진화 + Cross-validation
- **Forward Test**: 7/30/90/180일 시점 실현 알파 측정

### 8. DCF Valuator
- **모델**: 5년 FCF 추정 + Terminal Value
- **WACC**: CAPM (Beta + 무위험금리 + 시장프리미엄)
- **민감도**: WACC ±1%, 성장률 ±2%

### 9. Earnings Reviewer
- **데이터**: DART API (한국 공시) + yfinance (컨센서스)
- **추출**: 매출, 영업이익, 순이익 (당기/전년/2년)
- **점수**: YoY 비교 + 컨센서스 대비 어닝 서프라이즈

## 워크플로 동시성 제어

11개 워크플로가 동시에 `docs/` 파일에 쓸 때 race condition 방지:

```yaml
concurrency:
  group: docs-write
  cancel-in-progress: false
```

→ 같은 그룹에 속한 워크플로는 순차 실행, push 충돌 방지.

## 보안 설계

| 항목 | 메커니즘 |
|---|---|
| API 키 | GitHub Secrets only (커밋 금지) |
| `config/.env` | `.gitignore` 등록, 로컬 전용 |
| XML 파싱 | `defusedxml` (XXE 방어) |
| 외부 호출 | HTTPS only |
| subprocess | list args (shell injection 방지) |

## 대시보드 빌드 흐름

```
generate_dashboard_data.py
├── 종합요약_{date}.txt 파싱 (sectors, macro)
├── 매크로분석_{date}.xlsx 파싱 (FOMC, 자산전망)
├── 지정학리스크_{date}.xlsx 파싱
├── 포트폴리오_{date}.xlsx 파싱
├── docs/*.json 병합 (overseas, japan, bitcoin_standard, value, dcf, earnings)
├── commentary_engine 호출 (AI 코멘트)
├── 가격 최신화 (info.currentPrice, history NaN 회피)
└── docs/data.json 출력 → 대시보드 fetch
```

## 향후 개선 포인트

| 영역 | 우선순위 | 작업 |
|---|---|---|
| 모놀리스 분리 | 중 | 1500+ LOC 파일 (fed_macro, news_sentiment) 모듈화 |
| 로깅 통합 | 중 | print → logging.getLogger(__name__) |
| 예외 세분화 | 낮 | except Exception → 구체 예외 |
| 단위 테스트 | 낮 | 점수 계산 함수 위주 |
| 매직 넘버 | 낮 | config/constants.py 신설 |
