# AI 주식 분석 에이전트 시스템

Claude Desktop Cowork용 멀티 에이전트 주식 분석 팀입니다.

## 폴더 구조

```
ai-finance/
├── agents/                          # Cowork 에이전트 팀
│   ├── orchestrator.md              # 팀장 - 전체 워크플로우 지휘
│   ├── market-researcher.md         # 시장 리서처 - 뉴스/산업 동향
│   ├── financial-analyst.md         # 재무 분석가 - 재무제표 분석
│   ├── regime-analyst.md            # 레짐 분석가 - HMM 분석 해석
│   ├── macro-analyst.md             # 매크로 분석가 - 연준/거시경제 분석
│   ├── sentiment-analyst.md         # 심리 분석가 - 뉴스 심리 분석
│   ├── report-writer.md             # 리포트 작성자 - 종합 보고서
│   └── README.md                    # 이 파일
├── financial_analyzer.py            # 재무 데이터 수집 스크립트
├── hmm_regime_detector.py           # HMM 레짐 분석 스크립트
├── fed_macro_analyzer.py            # 연준/거시경제 데이터 수집 스크립트
├── news_sentiment_analyzer.py       # 뉴스 감성 분석 스크립트
├── portfolio_analyzer.py            # 포트폴리오 분석 스크립트
├── requirements.txt                 # Python 의존성
├── output/                          # 분석 결과 저장 폴더
│   ├── {ticker}_금융분석_{date}.xlsx
│   ├── {ticker}_HMM레짐분석_{date}.xlsx
│   ├── {ticker}_심리분석_{date}.xlsx
│   ├── 매크로분석_{date}.xlsx
│   ├── 포트폴리오분석_{date}.xlsx
│   ├── {ticker}_metrics.json
│   └── {ticker}_종합보고서_{date}.md
└── dashboard.html                   # 시각화 대시보드
```

## 사전 준비

### 1. Python 환경 설정

```bash
cd C:/Users/lacoi/Desktop/ai-finance/
pip install -r requirements.txt
```

### 2. 데이터 수집 테스트

```bash
# 단일 종목 전체 데이터
PYTHONIOENCODING=utf-8 python financial_analyzer.py --ticker 005930
PYTHONIOENCODING=utf-8 python hmm_regime_detector.py --ticker 005930
PYTHONIOENCODING=utf-8 python fed_macro_analyzer.py --ticker 005930
PYTHONIOENCODING=utf-8 python news_sentiment_analyzer.py --ticker 005930

# 포트폴리오 분석
PYTHONIOENCODING=utf-8 python portfolio_analyzer.py --tickers 005930,000660,035420 --weights 0.5,0.3,0.2

# 매크로 단독 분석
PYTHONIOENCODING=utf-8 python fed_macro_analyzer.py
```

`output/` 폴더에 Excel, JSON 파일이 생성되면 정상입니다.

## Cowork 연결 방법

1. **Claude Desktop** 실행
2. **Cowork** 메뉴 열기
3. 이 `agents/` 폴더를 Cowork 에이전트 디렉토리로 선택
4. 팀장(orchestrator)이 자동으로 서브 에이전트를 관리합니다

## 사용 예시

### 전체 종합 분석 (단일 종목)
```
삼성전자 종합 분석해줘
005930 분석 리포트 만들어줘
AAPL 전체 분석 부탁해
```

### 포트폴리오 분석
```
삼성전자 50%, SK하이닉스 30%, 네이버 20% 포트폴리오 분석해줘
AAPL 40%, MSFT 35%, NVDA 25% 포트폴리오 리포트 작성해줘
```

### 매크로 단독 분석
```
연준 정책 및 매크로 환경 분석해줘
거시경제 현황만 분석해줘
```

### 특정 분석만 요청
```
005930 레짐만 해석해줘
AAPL 재무 분석만 해줘
삼성전자 뉴스 심리 분석해줘
005930 매크로 영향 분석해줘
```

### 해외 주식
```
TSLA 종합 분석해줘
MSFT 재무 + 레짐 분석
NVDA 투자 보고서 작성해줘
```

## 에이전트 역할 요약

| 에이전트 | 역할 | 입력 | 출력 |
|----------|------|------|------|
| 팀장 | 워크플로우 지휘, 스크립트 실행 | 종목코드 | 진행 관리 |
| 시장 리서처 | 뉴스, 산업, 매크로 조사 | 종목명 | 시장 환경 분석 |
| 재무 분석가 | 재무제표 심층 분석 | Excel, JSON | 재무 분석 리포트 |
| 레짐 분석가 | HMM 결과 해석, 타이밍 | Excel | 레짐 분석 리포트 |
| 매크로 분석가 | 연준 정책, 경기 사이클 해석 | Excel | 매크로 환경 리포트 |
| 심리 분석가 | 뉴스 감성, 기술적 심리 분석 | Excel | 심리 분석 리포트 |
| 리포트 작성자 | 전체 결과 종합 | 모든 분석 결과 | 종합 투자 보고서 |

## 분석 워크플로우

### 단일 종목 분석

```
사용자: "005930 분석해줘"
        │
        v
[1] 팀장: 데이터 수집 (4개 Python 스크립트 실행)
        │
        v
[2] 병렬 실행:
    ├── 시장 리서처: 뉴스/산업 조사
    ├── 재무 분석가: Excel 데이터 분석
    ├── 레짐 분석가: HMM 결과 해석
    ├── 매크로 분석가: 거시경제 환경 해석
    └── 심리 분석가: 뉴스 감성 분석
        │
        v
[3] 리포트 작성자: 종합 보고서 작성
        │
        v
[4] 결과: output/005930_종합보고서_20260321.md
```

### 포트폴리오 분석

```
사용자: "삼성전자 50%, SK하이닉스 30%, 네이버 20% 포트폴리오 분석"
        │
        v
[1] 팀장: 개별 종목 + 포트폴리오 데이터 수집 (5개 스크립트)
        │
        v
[2] 병렬 에이전트 분석 (위와 동일)
        │
        v
[3] 리포트 작성자: 포트폴리오 시사점 포함 종합 보고서
```

### 매크로 단독 분석

```
사용자: "거시경제 환경 분석해줘"
        │
        v
[1] 팀장: fed_macro_analyzer.py 실행
        │
        v
[2] 매크로 분석가: 거시경제 환경 해석
        │
        v
[3] 리포트 작성자: 매크로 보고서 작성
```

## 주의사항

- Python 스크립트 실행 시 반드시 `PYTHONIOENCODING=utf-8` 환경변수 필요 (Windows)
- 인터넷 연결이 필요합니다 (주가 데이터 다운로드, 뉴스 검색)
- 분석 결과는 AI 기반 참고 자료이며, 투자 판단의 최종 책임은 투자자에게 있습니다
- 포트폴리오 분석 시 종목 수와 비중 합계(1.0)를 정확히 입력해야 합니다
