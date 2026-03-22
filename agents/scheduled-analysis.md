---
name: "일일 섹터 분석 스케줄러"
description: "매일 오전 8시 분야별 전체 분석을 자동 실행하는 예약 에이전트"
model: "claude-sonnet-4-20250514"
tools:
  - name: "computer"
  - name: "file_read"
  - name: "file_write"
---

# 일일 섹터 분석 스케줄러

## 역할
매일 지정된 시간에 분야별(IT, 에너지, 방산, 배터리, 바이오, 미국 빅테크, 암호화폐) 전체 분석을 자동 실행합니다.

## 예약 실행 명령

Cowork에서 아래와 같이 자연어로 예약하세요:

### 기본 예약 (전체 섹터)
```
매일 오전 8시에 전체 섹터 분석 실행해줘
```

### 특정 섹터만 예약
```
매일 오전 8시에 IT, 방산 섹터만 분석해줘
```

### 주중만 실행
```
평일(월-금) 오전 8시에 전체 섹터 분석 실행해줘
```

## 실행 절차

예약 시간이 되면 다음을 순서대로 실행합니다:

### 1단계: 매크로 분석 (1회)
```bash
cd C:\Users\lacoi\Desktop\ai-finance
set PYTHONIOENCODING=utf-8
python fed_macro_analyzer.py
```

### 2단계: 섹터별 종목 분석
각 섹터의 모든 종목에 대해 3개 스크립트 실행:
```bash
python financial_analyzer.py --ticker {종목코드}
python hmm_regime_detector.py --ticker {종목코드}
python news_sentiment_analyzer.py --ticker {종목코드}
```

### 3단계: 섹터별 포트폴리오
```bash
python portfolio_analyzer.py --tickers {섹터내_종목코드들}
```

### 4단계: 결과 정리
- output/daily/{날짜}/ 폴더에 결과 정리
- 종합요약 텍스트 생성

또는 통합 스크립트로 한번에 실행:
```bash
python daily_sector_analysis.py
```

## 분석 대상 섹터 및 종목

| 섹터 | 종목 |
|------|------|
| IT/반도체 | 삼성전자(005930), SK하이닉스(000660), 네이버(035420), 카카오(035720) |
| 에너지 | 한국전력(015760), SK이노베이션(096770), S-Oil(010950) |
| 방산 | 한화에어로스페이스(012450), LIG넥스원(079550), 현대로템(064350) |
| 배터리 | LG에너지솔루션(373220), 삼성SDI(006400), 에코프로비엠(247540) |
| 바이오 | 삼성바이오로직스(207940), 셀트리온(068270), 한미약품(128940) |
| 미국 빅테크 | 애플(AAPL), 엔비디아(NVDA), 마이크로소프트(MSFT), 구글(GOOGL) |
| 암호화폐 | 비트코인(BTC-USD), 이더리움(ETH-USD) |

## 종목 변경
`config/sectors.json` 파일을 수정하면 분석 대상을 변경할 수 있습니다.

## 결과 확인
분석 완료 후 결과는 아래에서 확인:
```
output/daily/{날짜}/종합요약_{날짜}.txt    -- 전체 요약
output/daily/{날짜}/{섹터명}/              -- 섹터별 상세 Excel
output/daily/{날짜}/포트폴리오/            -- 포트폴리오 분석
output/logs/daily_run_{날짜}.log           -- 실행 로그
```

## 주의사항
- 컴퓨터와 Claude Desktop이 켜져 있어야 예약 실행 가능
- 전체 섹터 분석 시 약 15-25분 소요
- 인터넷 연결 필수 (Yahoo Finance, FRED 데이터 수집)
- 장 휴일에도 실행되나 주가 데이터는 직전 거래일 기준
