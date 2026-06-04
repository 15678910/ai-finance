"""
국내 반도체 ETF 구성종목 비중 수집기
=====================================
pykrx의 get_etf_portfolio_deposit_file()을 이용해
주요 반도체 ETF 구성종목과 비중을 자동 수집합니다.

출력: docs/etf_holdings.json
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta, date

KST = timezone(timedelta(hours=9))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "docs", "etf_holdings.json")

# 주요 국내 반도체 ETF
ETF_LIST = [
    {"ticker": "396500", "name": "TIGER 반도체TOP10", "manager": "미래에셋", "aum_tril": 14.11, "color": "#22d3ee"},
    {"ticker": "091160", "name": "KODEX 반도체",      "manager": "삼성",    "aum_tril":  7.25, "color": "#4ade80"},
    {"ticker": "395270", "name": "HANARO K-반도체",   "manager": "NH아문디","aum_tril":  4.33, "color": "#f472b6"},
    {"ticker": "091230", "name": "TIGER 반도체",      "manager": "미래에셋","aum_tril":  1.64, "color": "#fbbf24"},
    {"ticker": "455850", "name": "SOL AI반도체소부장","manager": "신한",    "aum_tril":  1.00, "color": "#a78bfa"},
    {"ticker": "388420", "name": "RISE 비메모리반도체","manager": "KB",     "aum_tril":  0.35, "color": "#f87171"},
]


def fetch_holdings_pykrx(ticker: str, today_str: str) -> list:
    """pykrx로 ETF 구성종목 비중 수집."""
    try:
        from pykrx import stock as pkstock
        df = pkstock.get_etf_portfolio_deposit_file(ticker, today_str)
        if df is None or df.empty:
            return []
        results = []
        for _, row in df.iterrows():
            name = row.get("종목명", "") or row.get("Name", "")
            weight = row.get("비중", row.get("Weight", 0))
            if name and weight and float(weight) > 0.1:
                results.append({
                    "name": str(name),
                    "weight": round(float(weight), 2),
                    "ticker": str(row.get("티커", row.get("Ticker", ""))).zfill(6)
                            if row.get("티커") or row.get("Ticker") else "",
                })
        results.sort(key=lambda x: x["weight"], reverse=True)
        return results[:15]
    except Exception as e:
        print(f"  [WARN] pykrx 수집 실패 ({ticker}): {e}")
        return []


# 연구 기반 정적 데이터 (pykrx 실패 시 폴백, 2026-06-02 기준)
STATIC_HOLDINGS = {
    "396500": [  # TIGER 반도체TOP10
        {"name": "SK하이닉스",   "ticker": "000660", "weight": 33.62},
        {"name": "삼성전자",     "ticker": "005930", "weight": 26.04},
        {"name": "한미반도체",   "ticker": "042700", "weight": 12.04},
        {"name": "DB하이텍",     "ticker": "000990", "weight":  5.81},
        {"name": "리노공업",     "ticker": "058470", "weight":  4.82},
        {"name": "이오테크닉스", "ticker": "039030", "weight":  3.98},
        {"name": "원익IPS",      "ticker": "240810", "weight":  3.39},
        {"name": "HPSP",         "ticker": "403870", "weight":  2.38},
        {"name": "ISC",          "ticker": "095340", "weight":  2.13},
        {"name": "주성엔지니어링","ticker":"036930",  "weight":  1.85},
    ],
    "091160": [  # KODEX 반도체
        {"name": "SK하이닉스",   "ticker": "000660", "weight": 24.71},
        {"name": "삼성전자",     "ticker": "005930", "weight": 21.93},
        {"name": "한미반도체",   "ticker": "042700", "weight": 10.15},
        {"name": "리노공업",     "ticker": "058470", "weight":  4.51},
        {"name": "원익IPS",      "ticker": "240810", "weight":  3.43},
        {"name": "이오테크닉스", "ticker": "039030", "weight":  3.10},
        {"name": "DB하이텍",     "ticker": "000990", "weight":  2.30},
        {"name": "ISC",          "ticker": "095340", "weight":  2.28},
        {"name": "HPSP",         "ticker": "403870", "weight":  1.85},
        {"name": "주성엔지니어링","ticker":"036930",  "weight":  1.85},
    ],
    "395270": [  # HANARO K-반도체
        {"name": "SK하이닉스",   "ticker": "000660", "weight": 24.34},
        {"name": "삼성전자",     "ticker": "005930", "weight": 25.48},
        {"name": "삼성전기",     "ticker": "009150", "weight": 17.20},
        {"name": "한미반도체",   "ticker": "042700", "weight":  8.81},
        {"name": "리노공업",     "ticker": "058470", "weight":  3.77},
        {"name": "이오테크닉스", "ticker": "039030", "weight":  2.45},
        {"name": "DB하이텍",     "ticker": "000990", "weight":  2.54},
        {"name": "주성엔지니어링","ticker":"036930",  "weight":  2.28},
        {"name": "원익IPS",      "ticker": "240810", "weight":  2.03},
        {"name": "HPSP",         "ticker": "403870", "weight":  1.68},
    ],
    "091230": [  # TIGER 반도체
        {"name": "SK하이닉스",   "ticker": "000660", "weight": 33.19},
        {"name": "삼성전자",     "ticker": "005930", "weight": 21.83},
        {"name": "한미반도체",   "ticker": "042700", "weight":  6.63},
        {"name": "주성엔지니어링","ticker":"036930",  "weight":  3.53},
        {"name": "DB하이텍",     "ticker": "000990", "weight":  3.20},
        {"name": "리노공업",     "ticker": "058470", "weight":  2.65},
        {"name": "이오테크닉스", "ticker": "039030", "weight":  2.20},
        {"name": "파두",         "ticker": "440110", "weight":  2.17},
        {"name": "원익IPS",      "ticker": "240810", "weight":  1.93},
        {"name": "ISC",          "ticker": "095340", "weight":  1.80},
    ],
    "455850": [  # SOL AI반도체소부장
        {"name": "한미반도체",   "ticker": "042700", "weight": 12.50},
        {"name": "리노공업",     "ticker": "058470", "weight":  9.80},
        {"name": "이수페타시스", "ticker": "007660", "weight":  8.40},
        {"name": "원익IPS",      "ticker": "240810", "weight":  7.20},
        {"name": "이오테크닉스", "ticker": "039030", "weight":  6.80},
        {"name": "한솔케미칼",   "ticker": "014680", "weight":  6.50},
        {"name": "HPSP",         "ticker": "403870", "weight":  5.90},
        {"name": "주성엔지니어링","ticker":"036930",  "weight":  5.60},
        {"name": "솔브레인",     "ticker": "357780", "weight":  5.20},
        {"name": "ISC",          "ticker": "095340", "weight":  4.80},
    ],
    "388420": [  # RISE 비메모리반도체
        {"name": "삼성전자",     "ticker": "005930", "weight": 22.81},
        {"name": "SK하이닉스",   "ticker": "000660", "weight": 15.22},
        {"name": "삼성전기",     "ticker": "009150", "weight":  7.78},
        {"name": "SK스퀘어",     "ticker": "402340", "weight":  7.01},
        {"name": "DB하이텍",     "ticker": "000990", "weight":  6.69},
        {"name": "리노공업",     "ticker": "058470", "weight":  5.20},
        {"name": "한미반도체",   "ticker": "042700", "weight":  4.80},
        {"name": "이오테크닉스", "ticker": "039030", "weight":  4.50},
        {"name": "원익IPS",      "ticker": "240810", "weight":  3.90},
        {"name": "주성엔지니어링","ticker":"036930",  "weight":  3.20},
    ],
}


def main():
    today_str = date.today().strftime("%Y%m%d")
    print("=" * 55)
    print("  국내 반도체 ETF 구성종목 수집")
    print(f"  기준일: {today_str}")
    print("=" * 55)

    etf_data = []
    for etf in ETF_LIST:
        ticker = etf["ticker"]
        print(f"\n[{etf['name']}] ({ticker}) 수집 중...")
        holdings = fetch_holdings_pykrx(ticker, today_str)
        if not holdings:
            print(f"  → pykrx 실패, 정적 데이터 사용")
            holdings = STATIC_HOLDINGS.get(ticker, [])
        else:
            print(f"  → pykrx 성공: {len(holdings)}개 종목")

        if holdings:
            top = holdings[0]
            print(f"  1위: {top['name']} {top['weight']}%")

        etf_data.append({
            **etf,
            "holdings": holdings,
            "data_source": "pykrx" if holdings and holdings != STATIC_HOLDINGS.get(ticker, []) else "static",
        })

    # 종목별 ETF 노출도 계산 (어느 ETF에 얼마나 비중 있나)
    exposure = {}
    for etf in etf_data:
        for h in etf["holdings"]:
            name = h["name"]
            if name not in exposure:
                exposure[name] = {"ticker": h.get("ticker",""), "etfs": []}
            exposure[name]["etfs"].append({
                "etf_name": etf["name"],
                "etf_ticker": etf["ticker"],
                "weight": h["weight"],
                "aum_tril": etf["aum_tril"],
                "color": etf["color"],
            })
    # 각 종목의 가중 평균 노출도 계산 (AUM 가중)
    stock_exposure = []
    for name, data in exposure.items():
        total_aum = sum(e["aum_tril"] for e in data["etfs"])
        avg_weight = sum(e["weight"] * e["aum_tril"] for e in data["etfs"]) / total_aum if total_aum else 0
        etf_count = len(data["etfs"])
        stock_exposure.append({
            "name": name,
            "ticker": data["ticker"],
            "etf_count": etf_count,
            "avg_weight_aum": round(avg_weight, 2),
            "etfs": sorted(data["etfs"], key=lambda x: x["weight"], reverse=True),
        })
    stock_exposure.sort(key=lambda x: (x["etf_count"], x["avg_weight_aum"]), reverse=True)

    output = {
        "generated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "base_date": today_str,
        "etfs": etf_data,
        "stock_exposure": stock_exposure[:20],
    }

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] {OUTPUT_FILE} 저장 완료")
    print(f"  ETF {len(etf_data)}개 / 종목 노출도 {len(stock_exposure)}개")
    return 0


if __name__ == "__main__":
    sys.exit(main())
