"""
generate_dashboard_data.py
일일 분석 결과를 읽어 웹 대시보드용 docs/data.json을 생성합니다.

Usage:
    python generate_dashboard_data.py
    python generate_dashboard_data.py --date 20260322
    python generate_dashboard_data.py --date 20260322 --daily-dir output/daily/20260322
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from glob import glob

# ---------------------------------------------------------------------------
# 상수 / 기본값
# ---------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

FALLBACK_NAME_TO_TICKER = {
    "삼성전자": "005930",
    "SK하이닉스": "000660",
    "네이버": "035420",
    "카카오": "035720",
    "한화에어로스페이스": "012450",
    "LIG넥스원": "079550",
    "현대로템": "064350",
    "LG에너지솔루션": "373220",
    "삼성SDI": "006400",
    "에코프로비엠": "247540",
    "비트코인": "BTC-USD",
    "이더리움": "ETH-USD",
    "리플": "XRP-USD",
}

# sectors.json의 섹터 키 → 실제 서브디렉터리 이름 매핑
# sectors.json display_name 혹은 name 기준으로 매핑.
# 실제 output 디렉터리 이름은 슬래시 대신 언더스코어를 사용.
SECTOR_DIR_MAP = {
    "IT/반도체": "IT_반도체",
    "에너지": "에너지",
    "방산": "방산",
    "배터리/2차전지": "배터리_2차전지",
    "바이오/헬스케어": "바이오_헬스케어",
    "미국 빅테크": "미국_빅테크",
    "암호화폐": "암호화폐",
}

# ---------------------------------------------------------------------------
# 헬퍼: sectors.json 로드 → name_to_ticker 딕셔너리 생성
# ---------------------------------------------------------------------------

def load_name_to_ticker(config_path: str) -> dict:
    """sectors.json을 읽어 {한국어이름: ticker} 딕셔너리를 반환합니다."""
    mapping = {}
    try:
        with open(config_path, encoding="utf-8") as f:
            data = json.load(f)

        # 두 가지 구조 지원:
        # 1) {"sectors": {"KEY": {"tickers": {...}}}}  (명세서 형식)
        # 2) {"KEY": {"name": "...", "tickers": {...}}}  (실제 파일 형식)
        sectors_data = data.get("sectors", data)

        for sector_key, sector_val in sectors_data.items():
            tickers = sector_val.get("tickers", {})
            for ticker, kor_name in tickers.items():
                mapping[kor_name] = ticker

        print(f"[INFO] sectors.json 로드 완료: {len(mapping)}개 종목 매핑")
    except FileNotFoundError:
        print(f"[WARN] sectors.json 없음: {config_path} → fallback 매핑 사용")
    except Exception as e:
        print(f"[WARN] sectors.json 로드 실패 ({e}) → fallback 매핑 사용")

    if not mapping:
        mapping = FALLBACK_NAME_TO_TICKER.copy()

    return mapping


# ---------------------------------------------------------------------------
# 헬퍼: 종합요약 텍스트 파싱
# ---------------------------------------------------------------------------

def parse_summary_txt(filepath: str) -> dict:
    """종합요약_{date}.txt를 파싱해 매크로 정보와 섹터별 종목 데이터를 반환합니다."""
    result = {
        "date": "",
        "generated_at": "",
        "macro": {},
        "sectors": [],
    }

    try:
        with open(filepath, encoding="utf-8") as f:
            text = f.read()
    except FileNotFoundError:
        print(f"[ERROR] 종합요약 파일 없음: {filepath}")
        return result
    except Exception as e:
        print(f"[ERROR] 종합요약 파일 읽기 실패: {e}")
        return result

    lines = text.splitlines()

    # --- 날짜 파싱 (첫 줄) ---
    date_match = re.search(r"\((\d{4}-\d{2}-\d{2})\)", lines[0] if lines else "")
    if date_match:
        result["date"] = date_match.group(1)

    # --- 생성 시각 파싱 ---
    gen_match = re.search(r"\[생성 시각\]\s*(.+)", text)
    if gen_match:
        result["generated_at"] = gen_match.group(1).strip()

    # --- 매크로 환경 파싱 ---
    macro = {}
    cycle_m = re.search(r"경기 사이클:\s*(.+)", text)
    rate_m = re.search(r"금리:\s*([^\s(]+)(?:\s*\(FFR\s*([\d.]+%)\))?", text)
    infl_m = re.search(r"인플레이션:\s*([^\s(]+)(?:\s*\(CPI\s*([\d.]+%)\))?", text)

    if cycle_m:
        macro["cycle"] = cycle_m.group(1).strip()
    if rate_m:
        macro["rate"] = rate_m.group(1).strip()
        macro["ffr"] = rate_m.group(2).strip() if rate_m.group(2) else "N/A"
    if infl_m:
        macro["inflation"] = infl_m.group(1).strip()
        macro["cpi"] = infl_m.group(2).strip() if infl_m.group(2) else "N/A"

    result["macro"] = macro

    # --- 섹터별 요약 파싱 ---
    # 섹터 헤더: 줄 끝에 ':'가 오고 공백 없이 시작하는 줄
    # 종목 줄: "  종목이름 - 레짐: X / 심리: Y(점수)"
    sectors = []
    current_sector = None

    # [섹터별 요약] 이후 블록만 처리
    in_sector_block = False
    for line in lines:
        stripped = line.strip()

        if stripped == "[섹터별 요약]":
            in_sector_block = True
            continue

        if not in_sector_block:
            continue

        # 생성 시각 줄에서 블록 종료
        if stripped.startswith("[생성 시각]"):
            break

        # 빈 줄 무시
        if not stripped:
            continue

        # 섹터 헤더 패턴: "섹터명:"
        if re.match(r"^[^\s].+:$", stripped):
            sector_name = stripped.rstrip(":")
            current_sector = {"name": sector_name, "stocks": []}
            sectors.append(current_sector)
            continue

        # 종목 줄 패턴: "종목이름 - 레짐: X / 심리: Y(점수)"
        stock_m = re.match(
            r"^(.+?)\s*-\s*레짐:\s*(.+?)\s*/\s*심리:\s*(.+?)\(([+-]?[\d.]+)\)\s*$",
            stripped,
        )
        if stock_m and current_sector is not None:
            name = stock_m.group(1).strip()
            regime = stock_m.group(2).strip()
            sentiment_label = stock_m.group(3).strip()
            score_raw = stock_m.group(4).strip()
            # 점수 앞 부호 보존
            score_str = score_raw if score_raw.startswith(("+", "-")) else f"+{score_raw}"
            current_sector["stocks"].append(
                {
                    "name": name,
                    "regime": regime,
                    "sentiment_label": sentiment_label,
                    "sentiment_score": score_str,
                }
            )

    result["sectors"] = sectors
    return result


# ---------------------------------------------------------------------------
# 헬퍼: metrics.json 파싱
# ---------------------------------------------------------------------------

def load_metrics(metrics_path: str) -> dict:
    """티커_metrics.json 파일을 읽어 필요한 필드만 추출합니다."""
    defaults = {
        "price": None,
        "market_cap": None,
        "per": "N/A",
        "forward_per": None,
        "roe": None,
        "revenue_growth": None,
        "high_52w": None,
        "low_52w": None,
    }
    try:
        with open(metrics_path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[WARN] metrics.json 읽기 실패 ({metrics_path}): {e}")
        return defaults

    def _val(key, fallback=None):
        v = data.get(key, fallback)
        return v if v not in (None, "", "N/A") else fallback

    return {
        "price": _val("현재주가"),
        "market_cap": _val("시가총액(억)"),
        "per": data.get("PER", "N/A"),
        "forward_per": _val("Forward PER"),
        "roe": _val("ROE(%)"),
        "revenue_growth": _val("매출성장률(%)"),
        "high_52w": _val("52주최고"),
        "low_52w": _val("52주최저"),
    }


# ---------------------------------------------------------------------------
# 헬퍼: daily_dir에서 티커→metrics 경로 인덱스 빌드
# ---------------------------------------------------------------------------

def build_ticker_metrics_index(daily_dir: str) -> dict:
    """
    daily_dir 하위 모든 서브디렉터리를 스캔해 {ticker: metrics_path} 딕셔너리를 반환합니다.
    파일명 패턴: {ticker}_metrics.json
    """
    index = {}
    try:
        for entry in os.scandir(daily_dir):
            if not entry.is_dir():
                continue
            subdir = entry.path
            pattern = os.path.join(subdir, "*_metrics.json")
            for fpath in glob(pattern):
                fname = os.path.basename(fpath)
                ticker = fname.replace("_metrics.json", "")
                index[ticker] = fpath
    except Exception as e:
        print(f"[WARN] metrics 인덱스 빌드 실패: {e}")
    print(f"[INFO] metrics 인덱스: {len(index)}개 티커 발견")
    return index


# ---------------------------------------------------------------------------
# 메인 로직
# ---------------------------------------------------------------------------

def generate(date_str: str, daily_dir: str, output_path: str) -> bool:
    """데이터를 수집해 output_path에 JSON을 저장합니다."""

    print(f"[INFO] 날짜: {date_str}")
    print(f"[INFO] 일일 디렉터리: {daily_dir}")
    print(f"[INFO] 출력 경로: {output_path}")

    # 1) name → ticker 매핑 로드
    config_path = os.path.join(SCRIPT_DIR, "config", "sectors.json")
    name_to_ticker = load_name_to_ticker(config_path)

    # 2) 종합요약 파싱
    summary_filename = f"종합요약_{date_str}.txt"
    summary_path = os.path.join(daily_dir, summary_filename)
    print(f"[INFO] 종합요약 파싱 중: {summary_path}")
    parsed = parse_summary_txt(summary_path)

    if not parsed["date"]:
        parsed["date"] = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
    if not parsed["generated_at"]:
        parsed["generated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 3) metrics 인덱스 빌드
    ticker_to_metrics_path = build_ticker_metrics_index(daily_dir)

    # 4) 종목별 metrics 병합
    for sector in parsed["sectors"]:
        enriched_stocks = []
        for stock in sector["stocks"]:
            kor_name = stock["name"]
            ticker = name_to_ticker.get(kor_name)

            stock_entry = {
                "name": kor_name,
                "ticker": ticker or "UNKNOWN",
                "regime": stock["regime"],
                "sentiment_label": stock["sentiment_label"],
                "sentiment_score": stock["sentiment_score"],
                # metrics 기본값
                "price": None,
                "market_cap": None,
                "per": "N/A",
                "forward_per": None,
                "roe": None,
                "revenue_growth": None,
                "high_52w": None,
                "low_52w": None,
            }

            if ticker and ticker in ticker_to_metrics_path:
                metrics = load_metrics(ticker_to_metrics_path[ticker])
                stock_entry.update(metrics)
                print(f"[INFO]   {kor_name} ({ticker}): metrics 로드 완료")
            else:
                if ticker:
                    print(f"[WARN]   {kor_name} ({ticker}): metrics.json 없음")
                else:
                    print(f"[WARN]   {kor_name}: ticker 매핑 없음")

            enriched_stocks.append(stock_entry)
        sector["stocks"] = enriched_stocks

    # 5) 최종 JSON 구조 조립
    output_data = {
        "date": parsed["date"],
        "generated_at": parsed["generated_at"],
        "macro": parsed["macro"],
        "sectors": parsed["sectors"],
    }

    # 6) docs/ 디렉터리 생성 및 파일 저장
    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        print(f"[OK] docs/data.json 생성 완료: {output_path}")
        return True
    except Exception as e:
        print(f"[ERROR] JSON 저장 실패: {e}")
        return False


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args():
    today = datetime.now().strftime("%Y%m%d")
    parser = argparse.ArgumentParser(
        description="일일 분석 결과를 읽어 웹 대시보드용 docs/data.json을 생성합니다."
    )
    parser.add_argument(
        "--date",
        default=today,
        help=f"분석 날짜 (YYYYMMDD, 기본값: {today})",
    )
    parser.add_argument(
        "--daily-dir",
        default=None,
        help="일일 출력 디렉터리 경로 (기본값: output/daily/{date})",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    date_str = args.date

    if args.daily_dir:
        daily_dir = args.daily_dir
    else:
        daily_dir = os.path.join(SCRIPT_DIR, "output", "daily", date_str)

    output_path = os.path.join(SCRIPT_DIR, "docs", "data.json")

    success = generate(date_str, daily_dir, output_path)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
