# AI-Finance Sentinel — Multi-Asset Fusion Center 기획안

> **참조**: [d3mocide/Sovereign_Watch](https://github.com/d3mocide/Sovereign_Watch) (Multi-INT Fusion Center)
> **재해석**: 군사용 Multi-Intelligence → 금융용 Multi-Asset Fusion
> **코드네임**: AI-Finance Sentinel ("자본시장 파수꾼")

---

## 🎯 핵심 개념 매핑

| Sovereign Watch (군사) | AI-Finance Sentinel (금융) |
|---|---|
| Multi-INT Fusion | **Multi-Asset Fusion** |
| Pulse Architecture | **Financial Pulse Architecture** |
| Tiered AI Cognition | **Tiered Intelligence** (휴리스틱→통계→LLM) |
| Tactical Map View | **Tactical Capital Map** |
| ADS-B/AIS/위성/OSINT | **FRED/yfinance/KRX/Naver/RSS** |
| Sovereign Glass | **Finance Glass** (이미 부분 적용) |
| TAK Protocol | **Standard Asset Object Format** |
| Data Sovereignty | **Data Transparency** (출처 명시) |

---

## 🏗️ 1. 아키텍처 재구성

### A. 6개 Financial Pulse (독립 폴러)

각 Pulse는 독립적으로 작동, 표준화된 객체로 통합 버스(`data.json`)에 수렴.

```
┌─────────────────────────────────────────────────────────┐
│                  Asset Fusion Bus                        │
│              (docs/data.json — TAK-like)                 │
└─────────────────────────────────────────────────────────┘
        ▲          ▲          ▲          ▲          ▲          ▲
        │          │          │          │          │          │
   ┌────┴────┐ ┌──┴───┐ ┌────┴────┐ ┌──┴───┐ ┌────┴────┐ ┌──┴───┐
   │ MACRO   │ │ KR   │ │ CRYPTO  │ │ LEG  │ │ SEMI    │ │ NEWS │
   │ PULSE   │ │ PULSE│ │ PULSE   │ │ PULSE│ │ PULSE   │ │ PULSE│
   │ FRED    │ │Naver+│ │ CoinDesk│ │Congress│ Challenger│ Rss  │
   │ 신용    │ │yf+pyk│ │ Block   │ │+RSS  │ │ +RSS    │ │ 7-feeds│
   │ +금리   │ │+민감도│ │ +비트   │ │      │ │         │ │      │
   └─────────┘ └──────┘ └─────────┘ └──────┘ └─────────┘ └──────┘
   1d cron    1d+4h     4h cron     4h cron   4h cron     1h cron
```

| Pulse | 데이터 소스 | 갱신 주기 | 현재 구현 |
|---|---|---|---|
| 🏦 **MACRO PULSE** | FRED · BOJ · ECB | 1d (KST 07:00) | ✅ credit_spread_monitor.py |
| 🇰🇷 **KR EQUITY PULSE** | yfinance · Naver · pykrx | 1d 16:30 + 4h | ✅ value_screener · investor_flow |
| 🪙 **CRYPTO PULSE** | CoinDesk · The Block · yfinance | 4h | ✅ bitcoin_standard · clarity_act 일부 |
| 🏛️ **LEGISLATIVE PULSE** | Congress.gov · GovTrack · 3 RSS | 4h | ✅ clarity_act_monitor.py |
| 🔬 **SEMI PULSE** | yfinance · RSS · 챌린저 thesis | 4h | ✅ semi_challenger_monitor + sensitivity |
| 📰 **NEWS PULSE** | 8 RSS (Reuters/BBC/연합/Heisenberg/...) | 1h | ✅ breaking_news_monitor.py |

### B. Tiered Intelligence (계층화된 분석)

```
Asset Object → Intelligence Router → Output
                  │
                  ├─ Tier 1: 🔧 Heuristic Rules (즉시)
                  │  · 임계값 매칭 (HY > 5%p)
                  │  · 패턴 감지 (연속 매수/매도)
                  │  · 백분위 극단 (95%+ / 5%-)
                  │  현재 100% 구현 ✅
                  │
                  ├─ Tier 2: 📊 Statistical (분석)
                  │  · OLS 회귀 (반도체 민감도)
                  │  · IRR 계산 (DCF, Gordon)
                  │  · 다변량 분석 (KOSPI 베타)
                  │  현재 100% 구현 ✅
                  │
                  └─ Tier 3: 🧠 LLM (선택)
                     · 뉴스 요약·번역 (이미 Google Translate)
                     · 시나리오 합성 (향후 Claude API)
                     · 자연어 질의 (향후 Claude API)
                     번역만 구현 ⚠️ 확장 가능
```

---

## 🎨 2. UI 레이아웃 재구성

### 메인 화면 구조 (Tactical Capital View)

```
┌────────────────────────────────────────────────────────────────────────┐
│  AI-Finance Sentinel        🟢 6/6 Pulses Active    🔍 Cmd+K       ☰  │ ← 헤더
├────────────────────────────────────────────────────────────────────────┤
│ PULSE STATUS BAR                                                        │
│ 🟢 MACRO  🟢 KR-EQ  🟢 CRYPTO  🟢 LEG  🟢 SEMI  🟢 NEWS    📡 LIVE   │ ← 펄스 상태바
├──────────────┬─────────────────────────────────────────────────────────┤
│              │                                                          │
│ 🎯 FILTERS   │              TACTICAL CAPITAL MAP                       │
│ ─────────    │  ┌────────────────────────────────────────────────┐    │
│ Priority     │  │                                                │    │
│ ☑ P1 (3)     │  │     🌐 글로벌 자산 지도 (또는)                 │    │
│ ☑ P2 (12)    │  │     📊 섹터 트리맵                             │    │
│ ☑ P3 (28)    │  │     🕸️ 지식 그래프 (KG 확장)                  │    │
│              │  │     📈 시계열 히트맵                            │    │
│ Pulse        │  │                                                │    │
│ ☑ MACRO      │  └────────────────────────────────────────────────┘    │
│ ☑ KR EQUITY  │                                                          │
│ ☑ CRYPTO     │  ┌─────────────────────────────────────────────────┐   │
│ ☐ LEG        │  │ EVENT FEED (실시간 우선순위 큐)                  │   │
│ ☑ SEMI       │  ├─────────────────────────────────────────────────┤   │
│ ☑ NEWS       │  │ 🚨 P1  하이일드 OAS 임계 근접   FRED  →Act       │   │
│              │  │ ⚠️ P2  SK하이닉스 외인 14일    Naver →Watch      │   │
│ Markets      │  │ 🟡 P3  BBB 1년 최저           FRED  →Note        │   │
│ ☑ KOSPI      │  │ 🟢 P3  Cerebras 신규 뉴스     CoinDesk →Read    │   │
│ ☑ KOSDAQ     │  │ ... (43 more)                                   │   │
│ ☑ US Bond    │  └─────────────────────────────────────────────────┘   │
│ ☑ FX         │                                                          │
│              │                                                          │
│ Time Range   │                                                          │
│ ◉ Live       │                                                          │
│ ○ 1D         │                                                          │
│ ○ 1W         │                                                          │
│ ○ 1M         │                                                          │
│              │                                                          │
│ Layers       │                                                          │
│ ☑ Signals    │                                                          │
│ ☑ Sectors    │                                                          │
│ ☑ Flow       │                                                          │
│ ☐ Earnings   │                                                          │
│              │                                                          │
└──────────────┴─────────────────────────────────────────────────────────┘
```

### 핵심 UI 요소

#### 1. **Pulse Status Bar** (상단)
- 6개 Pulse 상태 한눈에 (🟢 정상 / 🟡 지연 / 🔴 실패)
- 각 Pulse 호버 시 마지막 갱신 시각·다음 cron 표시
- "LIVE" 표시 (실시간 갱신 중)

#### 2. **Filter Panel** (좌측, Sovereign Watch 스타일)
- **Priority**: P1 / P2 / P3 (자동 분류 + 카운트 배지)
- **Pulse**: 6개 펄스 토글
- **Markets**: KOSPI / KOSDAQ / US Bond / FX / Crypto
- **Time Range**: Live / 1D / 1W / 1M / 1Y
- **Layers**: Signals / Sectors / Flow / Earnings / Geopolitics

#### 3. **Tactical Capital Map** (중앙)
4가지 뷰 토글:
- 🌐 **글로벌 자산 지도** (geo 기반 — 한국·미국·일본·유럽 마커)
- 📊 **섹터 트리맵** (시총 비율 + 색상 = 등락률, Map of Market 스타일)
- 🕸️ **지식 그래프** (현재 KG 확장)
- 📈 **시계열 히트맵** (자산 × 시간 매트릭스)

#### 4. **Event Feed** (하단, Sovereign Watch 스타일)
- 모든 알림이 P1/P2/P3 우선순위로 정렬
- 호버 시 액션 버튼 (→Act / →Watch / →Note / →Read)
- 클릭 시 우측 Inspector Panel 슬라이드 인

#### 5. **Object Inspector Panel** (우측, 슬라이드)
클릭한 객체의 모든 정보 통합:
```
┌──────────────────────┐
│ × INSPECTOR          │
├──────────────────────┤
│ 📈 SK하이닉스         │
│ 000660 · KOSPI · P2  │
├──────────────────────┤
│ 📡 DATA SOURCES      │
│ yfinance · Naver     │
│ 마지막: 17:43 KST    │
├──────────────────────┤
│ 🎯 SNAPSHOT          │
│ 가격 159,000 +2.3%   │
│ IRR 12.5% MoS +3.2pt│
│ NQ β +1.60 R²0.15   │
│ SOX β +0.86         │
├──────────────────────┤
│ 🚨 ACTIVE SIGNALS    │
│ ⚠️ 외인 14일 매도    │
│ 🟢 IRR 매수 검토     │
├──────────────────────┤
│ 🔗 RELATED           │
│ → 삼성전자 (corr)    │
│ → SOX 지수 (β)       │
│ → Cerebras (위협)    │
├──────────────────────┤
│ [→ Full Analysis]    │
│ [⭐ Watch] [📝 Note] │
└──────────────────────┘
```

---

## 🎨 3. 디자인 시스템: "Finance Glass"

### 색상 팔레트 (Sovereign Glass 스타일 + 금융 조정)

```
배경:
  --bg-primary:   #0a0e17  (현재 사용 중)
  --bg-secondary: #0f1620
  --bg-glass:     rgba(15, 23, 42, 0.6)

액센트 (Pulse 별 컬러):
  --macro:    #ef4444  (빨강 - 거시·신용)
  --kr-eq:    #22c55e  (녹색 - 한국 종목)
  --crypto:   #f97316  (주황 - 암호화폐)
  --leg:      #fb923c  (오렌지 - 입법)
  --semi:     #06b6d4  (시안 - 반도체)
  --news:     #9ca3af  (회색 - 뉴스)

우선순위:
  --p1:       #dc2626  (P1 Critical)
  --p2:       #f97316  (P2 Alert)
  --p3:       #fbbf24  (P3 Watch)
  --p4:       #94a3b8  (P4 Info)

상태:
  --pulse-live:   #4ade80  (🟢 LIVE)
  --pulse-delay:  #facc15  (🟡 지연)
  --pulse-down:   #f87171  (🔴 실패)
```

### 타이포그래피
- **Brand**: Inter / Pretendard (sans-serif)
- **Data**: JetBrains Mono / D2Coding (monospace) — 데이터/티커
- **UI**: System Font Stack

### 컴포넌트 라이브러리 (재사용)
```
.pulse-badge     상태 점 + 라벨
.priority-tag    P1/P2/P3 배지
.source-tag      데이터 출처 (FRED/yfinance/etc.)
.event-card      Event Feed 카드
.inspector-row   Inspector 한 행
.layer-toggle    레이어 체크박스
.filter-chip     필터 칩
.tactical-tab    맵 뷰 전환 탭
```

---

## 📦 4. 데이터 모델: Standard Asset Object (SAO)

Sovereign Watch의 TAK Protocol → 우리만의 SAO 정의:

```typescript
interface AssetObject {
  // 식별
  id: string;              // "stock:005930" / "fred:DGS10" / "bill:HR3633"
  type: 'Stock' | 'Index' | 'Currency' | 'Bond' | 'Bill' | 'NewsItem' | 'MacroEvent';
  name: string;
  
  // 분류
  pulse: 'MACRO' | 'KR_EQUITY' | 'CRYPTO' | 'LEGISLATIVE' | 'SEMI' | 'NEWS';
  market?: 'KOSPI' | 'KOSDAQ' | 'NYSE' | 'NASDAQ' | 'KRX_DEBT';
  sector?: string;
  
  // 메타데이터
  source: string[];        // ["FRED", "yfinance"]
  last_updated: string;    // ISO 8601
  
  // 핵심 값
  value: number;
  unit: string;
  
  // 우선순위 (자동 계산)
  priority: 'P1' | 'P2' | 'P3' | 'P4';
  signals: Signal[];       // 활성 시그널 목록
  
  // 시각화용
  position?: { x: number; y: number; lat?: number; lng?: number; };
  size?: number;           // 시각화 크기 (market cap 등)
  color?: string;          // 등락률 색상 등
  
  // 메트릭
  metrics: Record<string, any>;
  
  // 관계 (KG 호환)
  relations: { type: string; target_id: string; weight?: number; data?: any }[];
}

interface Signal {
  type: string;
  strength: 1 | 2 | 3 | 4 | 5;
  emoji: string;
  label: string;
  description: string;
  detected_at: string;
}
```

→ 모든 모듈이 이 표준 형식으로 데이터 생성 시 통합 일관성 확보.

---

## 🛣️ 5. 단계별 구현 로드맵

### Phase α — Foundation (Week 1, ~12시간)
- [ ] `core/sao.py` — Standard Asset Object 정의 + 기존 모듈 어댑터
- [ ] **Pulse Status Bar** 헤더 컴포넌트
- [ ] **Event Feed** 통합 알림 큐 (현재 분산된 시그널 통합)
- [ ] **Filter Panel** 좌측 사이드바 (현재 Claude Docs 사이드바 확장)
- [ ] 우선순위 자동 계산 (P1/P2/P3/P4)

→ **이미 있는 데이터를 한 화면에 통합**. 가장 큰 가치.

### Phase β — Tactical Map (Week 2, ~10시간)
- [ ] **섹터 트리맵** (D3.js treemap) — 시총 × 등락률
- [ ] **글로벌 자산 지도** (MapLibre — 가벼움) — 한국/미국/일본/유럽 마커
- [ ] **시계열 히트맵** (D3.js) — 자산 × 시간
- [ ] 4가지 뷰 토글 (탭 전환)
- [ ] **레이어 토글** (Signals/Sectors/Flow/Earnings)

### Phase γ — Inspector + Search (Week 3, ~8시간)
- [ ] **Object Inspector Panel** (우측 슬라이드, KG 확장)
- [ ] **Cmd+K Command Palette**
- [ ] **Action Affordances** (호버 시 → Act/Watch/Note/Read)

### Phase δ — Live + AI (Week 4+, ~12시간)
- [ ] **WebSocket Live Updates** (현재 cron → 실시간)
- [ ] **Tier 3 AI Router** (Claude API 통합, 옵션)
- [ ] **자연어 질의** ("외국인 5일 이상 매도 종목" → SQL/필터 자동)

---

## 💰 6. 비용·기술적 고려

### 모두 무료 가능
- D3.js / MapLibre GL JS / vis.js — 모두 무료 오픈소스
- 추가 인프라 불필요 — 기존 GH Actions + Pages
- WebSocket은 GH Pages 미지원이라 폴링 유지 (1분 갱신 가능)

### 선택 비용 (있을 시)
- Mapbox 토큰: 무료 한도 50K req/month (충분)
- Claude/GPT API: 월 $1~5 (LLM 추론 옵션)

---

## 🎯 7. Sovereign Watch vs AI-Finance Sentinel 핵심 차이

| 항목 | Sovereign Watch | AI-Finance Sentinel |
|---|---|---|
| 도메인 | 군사·국방 정보 | 금융·자본시장 |
| 데이터 주체 | 항공기/선박/위성 | 종목/지수/통화/채권 |
| 우선순위 의미 | 위협 등급 | 행동 시급도 (Act/Watch/Note/Read) |
| 지도 의미 | 지리적 위치 | 시장·섹터·관계 위치 |
| 실시간 요구 | 초·분 단위 | 시간·일 단위 (시장 특성) |
| Tiered AI | Llama3 → Claude | 휴리스틱 → 회귀 → LLM |
| 데이터 주권 | 셀프 호스팅 (보안) | 데이터 투명성 (출처 명시) |
| 사용자 | 정보 분석가 | 투자자·트레이더 |

---

## 🚦 8. 즉시 실행 가능한 다음 단계

1. **🎯 Phase α 시작 신호 받으면**:
   - `core/sao.py` 데이터 표준 정의 작성
   - 기존 모듈 JSON 출력을 SAO 형식으로 매핑
   - `Pulse Status Bar` HTML/CSS 추가
   - `Event Feed` 통합 큐 컴포넌트

2. **혹은 특정 부분만 선택**:
   - "Pulse Status Bar만 먼저" → 2시간
   - "Event Feed만 먼저" → 4시간
   - "섹터 트리맵만 먼저" → 3시간

3. **컨셉 검증**:
   - 와이어프레임 먼저 (위 ASCII 레이아웃)
   - 1개 화면만 Figma/HTML 프로토타입

---

## 💡 핵심 인사이트

**Sovereign Watch가 잘 하는 것**:
- 다양한 데이터 소스를 **하나의 운영 화면**에 통합
- **우선순위 + 큐**로 인지 부담 감소
- **레이어 시스템**으로 정보 밀도 조절
- **객체 중심** 탐색 (지도 클릭 → 상세)

**우리가 더 잘할 수 있는 것**:
- **금융 도메인 특화 지표** (IRR, OAS, β 등)
- **한국 시장 깊이** (외국인·기관 매매, 환율, 입법)
- **이미 구축된 모듈 수** (12+ 분석기)
- **무료 운영** (군사용 인프라 비용 X)

→ **AI-Finance Sentinel**은 "투자자용 Sovereign Watch"가 될 수 있음.

---

## 🎬 결론

**Sovereign Watch의 가장 가치 있는 패턴**:
1. ✅ **Pulse Architecture** (데이터 소스 독립화) — 이미 우리도 모듈 구조
2. ⭐ **통합 Event Feed** (모든 알림 한 곳) — 가장 큰 미구현 가치
3. ⭐ **Filter/Layer System** (필터링·레이어) — 정보 조절
4. ✅ **Tiered AI** (계층화) — 우리도 이미 있음
5. ⭐ **Standard Object Format** (TAK 같은 통합 형식) — 향후 확장성

**최우선 적용 권고**: Phase α (Foundation 12시간)
- 통합 Event Feed가 매일 사용자 경험에 가장 큰 차이.
- 현재 8개 섹션 흩어진 알림을 하나의 우선순위 큐로.

진행 의사 있으시면 알려주세요. 또는 특정 부분(예: Event Feed만, Treemap만)부터 시작 가능.
