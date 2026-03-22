# AI 금융 분석 시스템 — 사용 가이드

## 📁 파일 구성
```
financial_analyzer.py   ← 메인 Python 스크립트
requirements.txt        ← 필요 라이브러리
dashboard.html          ← 시스템 전체 설계도 (브라우저로 열기)
output/                 ← 생성된 Excel/JSON 파일 저장 폴더
```

---

## 🚀 시작하기

### 1단계: 라이브러리 설치
```bash
pip install -r requirements.txt
```

### 2단계: 데이터 수집 실행
```bash
# 국내 주식 (삼성전자)
python financial_analyzer.py --ticker 005930

# 미국 주식 (애플)
python financial_analyzer.py --ticker AAPL

# 기간 지정 (5년)
python financial_analyzer.py --ticker 005930 --period 5y
```

### 3단계: Excel 파일을 Cowork에 업로드
생성된 `output/005930_금융분석_YYYYMMDD.xlsx` 파일을
Claude Desktop → Cowork에 첨부하고 아래 명령어 사용:

```
/dcf-model 이 파일로 DCF 가치평가 모델 구축해줘
/comps-analysis 동종사 비교 분석해줘  
/3-statement-model 통합 재무모델 만들어줘
```

---

## 🔑 DART API 키 설정 (한국 주식 공시 수집용)

1. https://opendart.fss.or.kr/ 회원가입
2. API 키 발급 (무료)
3. 환경변수 설정:

**Windows:**
```cmd
setx DART_API_KEY "your_api_key_here"
```

**Mac/Linux:**
```bash
export DART_API_KEY="your_api_key_here"
```

---

## 📊 지원 데이터
- 주가 (최대 5년)
- 손익계산서, 대차대조표, 현금흐름표
- PER, PBR, EV/EBITDA, ROE, ROA 등 40+ 지표
- DART 공시 목록 (API 키 필요)

---

## 🔄 전체 워크플로우
```
[Claude Code]          python financial_analyzer.py --ticker 005930
      ↓ (Excel 생성)
[Cowork]               /dcf-model /comps-analysis /3-statement-model
      ↓ (보고서 생성)
[Chat]                 결과 해석, 투자 의사결정 토론
```
