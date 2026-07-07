// naver-stock-proxy — 네이버 통합시세(NXT 포함) CORS 프록시 (Cloudflare Workers)
// =============================================================================
// 목적: GitHub Pages(정적)에서 브라우저가 네이버 시세를 직접 못 부르는 CORS 문제를
//       우회하기 위해, 이 Worker가 서버 사이드로 네이버를 대신 호출해 CORS 헤더를 붙여 반환.
//
// 배포 방법:
//   1) https://dash.cloudflare.com → Workers & Pages → Create → Worker(이름 예: naver-stock)
//   2) "Edit code"에서 이 파일 내용을 전부 붙여넣고 → Deploy
//   3) 배포 URL(예: https://naver-stock.<계정>.workers.dev)을 복사
//   4) docs/index.html 의 PF_LIVE_PROXY 상수에 그 URL을 넣기
//
// 사용: GET https://<worker>.workers.dev/?codes=000660,108490
//   응답: {"000660":{"price":2343000,"name":"SK하이닉스","rate":-2.56,"session":"PRE_MARKET",...}, ...}
//
// 시세 선택 로직:
//   · 정규장(09:00~15:30): closePrice(실시간 체결가)
//   · NXT 프리마켓(08:00~)/애프터마켓(~20:00) 거래 중: overMarketPriceInfo.overPrice
// =============================================================================

const SRC = 'https://polling.finance.naver.com/api/realtime/domestic/stock/';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': '*',
  'Cache-Control': 'no-store',
  'Content-Type': 'application/json; charset=utf-8',
};

const toNum = (s) => {
  if (s == null) return null;
  const n = Number(String(s).replace(/[,\s]/g, ''));
  return Number.isFinite(n) ? n : null;
};

export default {
  async fetch(request) {
    if (request.method === 'OPTIONS') return new Response(null, { headers: CORS });
    const url = new URL(request.url);
    const codes = (url.searchParams.get('codes') || '').replace(/[^0-9,]/g, '');
    if (!codes) {
      return new Response(JSON.stringify({ error: 'codes 파라미터 필요 (예: ?codes=000660,108490)' }), { status: 400, headers: CORS });
    }
    try {
      const r = await fetch(SRC + codes, {
        headers: { 'User-Agent': 'Mozilla/5.0', 'Referer': 'https://finance.naver.com/', 'Accept': 'application/json' },
      });
      if (!r.ok) return new Response(JSON.stringify({ error: 'naver ' + r.status }), { status: 502, headers: CORS });
      const j = await r.json();
      const datas = (j && j.datas) || [];
      const out = {};
      for (const d of datas) {
        const code = d.itemCode || d.cd;
        if (!code) continue;
        const over = d.overMarketPriceInfo;
        let price = null, session = 'REGULAR', rate = toNum(d.fluctuationsRatio);
        // NXT 프리/애프터마켓이 거래 중(OPEN)이고 유효 시세가 있으면 그 값을, 아니면 정규장 현재가
        if (over && over.overMarketStatus === 'OPEN' && over.overPrice && over.overPrice !== '-') {
          price = toNum(over.overPrice);
          session = over.tradingSessionType || 'OVER';
          rate = toNum(over.fluctuationsRatio);
        } else {
          price = toNum(d.closePrice);
        }
        out[code] = {
          price,
          name: d.stockName || code,
          rate,
          session,
          market: d.marketStatus || null,
          at: d.localTradedAt || null,
        };
      }
      return new Response(JSON.stringify(out), { headers: CORS });
    } catch (e) {
      return new Response(JSON.stringify({ error: String((e && e.message) || e) }), { status: 502, headers: CORS });
    }
  },
};
