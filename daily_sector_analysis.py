"""
일일 섹터별 종합 분석 시스템 - 마스터 오케스트레이션 스크립트
===========================================================
사용법:
  python daily_sector_analysis.py                        # 전체 섹터 분석
  python daily_sector_analysis.py --sectors IT,DEFENSE   # 특정 섹터만
  python daily_sector_analysis.py --skip-macro           # 매크로 분석 건너뛰기
  python daily_sector_analysis.py --skip-portfolio       # 포트폴리오 건너뛰기
  python daily_sector_analysis.py --reset-config         # 섹터 설정 초기화
  python daily_sector_analysis.py --dry-run              # 실행 계획만 출력

생성 구조:
  output/daily/{날짜}/macro/
  output/daily/{날짜}/{섹터명}/
  output/daily/{날짜}/포트폴리오/
  output/daily/{날짜}/종합요약_{날짜}.txt
  output/daily/{날짜}/errors.log
"""

import os
import sys
import json
import time
import glob
import shutil
import argparse
import logging
import subprocess
from datetime import datetime
from pathlib import Path


# ====================================================================
# 0. 기본 상수
# ====================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
CONFIG_FILE = os.path.join(CONFIG_DIR, "sectors.json")
OUTPUT_BASE = os.path.join(BASE_DIR, "output")
SCRIPT_OUTPUT_DIR = os.path.join(BASE_DIR, "output")

DEFAULT_SECTORS = {
    "IT": {
        "name": "IT/반도체",
        "tickers": {
            "005930": "삼성전자",
            "000660": "SK하이닉스",
            "035420": "네이버",
            "035720": "카카오",
        }
    },
    "ENERGY": {
        "name": "에너지",
        "tickers": {
            "015760": "한국전력",
            "096770": "SK이노베이션",
            "010950": "S-Oil",
        }
    },
    "DEFENSE": {
        "name": "방산",
        "tickers": {
            "012450": "한화에어로스페이스",
            "079550": "LIG넥스원",
            "064350": "현대로템",
        }
    },
    "BATTERY": {
        "name": "배터리/2차전지",
        "tickers": {
            "373220": "LG에너지솔루션",
            "006400": "삼성SDI",
            "247540": "에코프로비엠",
        }
    },
    "BIO": {
        "name": "바이오/헬스케어",
        "tickers": {
            "207940": "삼성바이오로직스",
            "068270": "셀트리온",
            "128940": "한미약품",
        }
    },
    "US_TECH": {
        "name": "미국 빅테크",
        "tickers": {
            "AAPL": "애플",
            "NVDA": "엔비디아",
            "MSFT": "마이크로소프트",
            "GOOGL": "구글",
        }
    },
    "CRYPTO": {
        "name": "암호화폐",
        "tickers": {
            "BTC-USD": "비트코인",
            "ETH-USD": "이더리움",
        }
    },
}

# 스크립트 경로
SCRIPTS = {
    "financial":  os.path.join(BASE_DIR, "financial_analyzer.py"),
    "hmm":        os.path.join(BASE_DIR, "hmm_regime_detector.py"),
    "macro":      os.path.join(BASE_DIR, "fed_macro_analyzer.py"),
    "sentiment":  os.path.join(BASE_DIR, "news_sentiment_analyzer.py"),
    "portfolio":  os.path.join(BASE_DIR, "portfolio_analyzer.py"),
    "geopolitical": os.path.join(BASE_DIR, "geopolitical_analyzer.py"),
}

# 스크립트별 출력 파일 패턴 (output/ 기준 상대경로)
# {ticker} 와 {date} 는 런타임에 치환
FILE_PATTERNS = {
    "financial":  "{ticker}_금융분석_{date}.xlsx",
    "hmm":        "{ticker}_HMM레짐분석_{date}.xlsx",
    "sentiment":  "{ticker}_심리분석_{date}.xlsx",
    "macro":      "매크로분석_{date}.xlsx",
    "geopolitical": "지정학리스크_{date}.xlsx",
    "portfolio":  "포트폴리오분석_{date}.xlsx",
    # financial_analyzer 가 추가 생성하는 파일
    "metrics":    "{ticker}_metrics.json",
}


# ====================================================================
# 1. 설정 파일 관리
# ====================================================================
def ensure_config(reset=False):
    """섹터 설정 파일 로드. 없으면 생성, reset=True면 초기화."""
    Path(CONFIG_DIR).mkdir(parents=True, exist_ok=True)

    if reset or not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_SECTORS, f, ensure_ascii=False, indent=2)
        if reset:
            print(f"  [설정] 섹터 설정이 초기화되었습니다: {CONFIG_FILE}")
        else:
            print(f"  [설정] 기본 섹터 설정 파일 생성: {CONFIG_FILE}")
        return DEFAULT_SECTORS

    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            sectors = json.load(f)
        print(f"  [설정] 섹터 설정 로드 완료: {CONFIG_FILE}")
        return sectors
    except (json.JSONDecodeError, IOError) as e:
        print(f"  [경고] 설정 파일 읽기 실패 ({e}). 기본값으로 진행합니다.")
        return DEFAULT_SECTORS


# ====================================================================
# 2. 출력 디렉토리 생성
# ====================================================================
def create_output_dirs(date_str, sectors, skip_macro=False, skip_portfolio=False):
    """날짜별 출력 디렉토리 구조 생성."""
    daily_dir = os.path.join(OUTPUT_BASE, "daily", date_str)

    if not skip_macro:
        Path(os.path.join(daily_dir, "macro")).mkdir(parents=True, exist_ok=True)

    for sector_key, sector_info in sectors.items():
        sector_name = sector_info["name"]
        # 폴더명에서 / 를 _ 로 변환
        safe_name = sector_name.replace("/", "_")
        Path(os.path.join(daily_dir, safe_name)).mkdir(parents=True, exist_ok=True)

    if not skip_portfolio:
        Path(os.path.join(daily_dir, "포트폴리오")).mkdir(parents=True, exist_ok=True)

    return daily_dir


# ====================================================================
# 3. 서브프로세스 실행
# ====================================================================
def run_script(script_path, args_list, description=""):
    """스크립트를 서브프로세스로 실행. 결과를 반환."""
    cmd = [sys.executable, script_path] + args_list

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    start = time.time()
    try:
        result = subprocess.run(
            cmd,
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            env=env,
            timeout=600,  # 10분 타임아웃
        )
        elapsed = time.time() - start
        return {
            "success": result.returncode == 0,
            "returncode": result.returncode,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "elapsed": elapsed,
            "description": description,
        }
    except subprocess.TimeoutExpired:
        elapsed = time.time() - start
        return {
            "success": False,
            "returncode": -1,
            "stdout": "",
            "stderr": "타임아웃 (10분 초과)",
            "elapsed": elapsed,
            "description": description,
        }
    except Exception as e:
        elapsed = time.time() - start
        return {
            "success": False,
            "returncode": -1,
            "stdout": "",
            "stderr": str(e),
            "elapsed": elapsed,
            "description": description,
        }


def move_output_file(filename, dest_dir):
    """output/ 에서 생성된 파일을 대상 디렉토리로 이동."""
    src = os.path.join(SCRIPT_OUTPUT_DIR, filename)
    if os.path.exists(src):
        dst = os.path.join(dest_dir, filename)
        # 동일 파일명이 이미 있으면 덮어쓰기
        if os.path.exists(dst):
            os.remove(dst)
        shutil.move(src, dst)
        return dst
    return None


def find_and_move_output(pattern, dest_dir):
    """glob 패턴으로 output/ 에서 파일을 찾아 이동. 이동된 파일 경로 리스트 반환."""
    search_pattern = os.path.join(SCRIPT_OUTPUT_DIR, pattern)
    found = glob.glob(search_pattern)
    moved = []
    for src in found:
        filename = os.path.basename(src)
        dst = os.path.join(dest_dir, filename)
        if os.path.exists(dst):
            os.remove(dst)
        shutil.move(src, dst)
        moved.append(dst)
    return moved


def format_elapsed(seconds):
    """초를 읽기 좋은 형태로 변환."""
    if seconds < 60:
        return f"{seconds:.0f}초"
    minutes = int(seconds // 60)
    secs = int(seconds % 60)
    return f"{minutes}분 {secs}초"


# ====================================================================
# 4. 분석 결과 파싱 (stdout에서 핵심 정보 추출)
# ====================================================================
def parse_macro_stdout(stdout):
    """매크로 분석 stdout에서 요약 정보 추출."""
    info = {
        "cycle_stage": "N/A",
        "rate_cycle": "N/A",
        "inflation_status": "N/A",
        "ffr": "N/A",
        "cpi": "N/A",
    }
    for line in stdout.splitlines():
        stripped = line.strip()
        if "경기 사이클" in stripped or "현재 단계" in stripped:
            # [확장기] 형태 추출
            start = stripped.find('[')
            end = stripped.find(']')
            if start != -1 and end != -1:
                info["cycle_stage"] = stripped[start + 1:end]
        if "금리 사이클" in stripped:
            start = stripped.find('[')
            end = stripped.find(']')
            if start != -1 and end != -1:
                info["rate_cycle"] = stripped[start + 1:end]
        if "상태:" in stripped and ("인플" in stripped or "CPI" in stripped):
            # 상태: 적정 형태
            parts = stripped.split("상태:")
            if len(parts) > 1:
                info["inflation_status"] = parts[1].strip()
        if "기준금리" in stripped and "FFR" in stripped:
            # 값 추출 시도
            for part in stripped.split():
                try:
                    val = float(part.replace('%', ''))
                    info["ffr"] = f"{val:.2f}%"
                    break
                except ValueError:
                    continue
        if "CPI YoY" in stripped:
            for part in stripped.split():
                try:
                    val = float(part.replace('%', ''))
                    info["cpi"] = f"{val:.2f}%"
                    break
                except ValueError:
                    continue
    return info


def parse_regime_stdout(stdout):
    """HMM 레짐 분석 stdout에서 현재 레짐 추출."""
    regime = "N/A"
    for line in stdout.splitlines():
        stripped = line.strip()
        if "현재 레짐:" in stripped:
            start = stripped.find('[')
            end = stripped.find(']')
            if start != -1 and end != -1:
                regime = stripped[start + 1:end]
            break
    return regime


def parse_sentiment_stdout(stdout):
    """심리 분석 stdout에서 종합 심리 추출."""
    label = "N/A"
    score = "N/A"
    for line in stdout.splitlines():
        stripped = line.strip()
        if "종합 심리:" in stripped:
            start = stripped.find('[')
            end = stripped.find(']')
            if start != -1 and end != -1:
                label = stripped[start + 1:end]
            # 점수 추출
            paren_start = stripped.find('(', end)
            paren_end = stripped.find('점', paren_start)
            if paren_start != -1 and paren_end != -1:
                score = stripped[paren_start + 1:paren_end].strip()
            break
    return label, score


def parse_portfolio_stdout(stdout):
    """포트폴리오 분석 stdout에서 Sharpe 비율 추출."""
    sharpe = "N/A"
    for line in stdout.splitlines():
        stripped = line.strip()
        if "Sharpe" in stripped or "샤프" in stripped:
            for part in stripped.split():
                try:
                    val = float(part)
                    sharpe = f"{val:.2f}"
                    break
                except ValueError:
                    continue
            if sharpe != "N/A":
                break
    return sharpe


# ====================================================================
# 5. 에러 로깅
# ====================================================================
class ErrorLogger:
    """분석 실행 중 발생한 에러를 수집하고 파일로 저장."""

    def __init__(self):
        self.errors = []

    def add(self, category, ticker, script, message, stderr=""):
        self.errors.append({
            "time": datetime.now().strftime("%H:%M:%S"),
            "category": category,
            "ticker": ticker,
            "script": script,
            "message": message,
            "stderr": stderr,
        })

    def save(self, filepath):
        if not self.errors:
            return
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"=== 분석 에러 로그 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===\n\n")
            for i, err in enumerate(self.errors, 1):
                f.write(f"[{i}] {err['time']} | {err['category']} | {err['ticker']} | {err['script']}\n")
                f.write(f"    오류: {err['message']}\n")
                if err['stderr']:
                    # stderr 첫 5줄만
                    stderr_lines = err['stderr'].strip().splitlines()[:5]
                    for line in stderr_lines:
                        f.write(f"    > {line}\n")
                f.write("\n")

    def print_summary(self):
        if not self.errors:
            print("\n  오류 없음 - 전체 분석 정상 완료")
            return
        print(f"\n  [경고] {len(self.errors)}건의 오류 발생:")
        for err in self.errors:
            ticker_str = f" ({err['ticker']})" if err['ticker'] else ""
            print(f"    - {err['category']}{ticker_str}: {err['message']}")


# ====================================================================
# 6. 드라이런 (실행 계획 출력)
# ====================================================================
def print_dry_run(sectors, skip_macro, skip_portfolio, date_str):
    """실행할 작업 목록만 출력하고 종료."""
    print("\n" + "=" * 60)
    print("  [DRY RUN] 실행 계획")
    print("=" * 60)
    print(f"  분석 날짜: {date_str}")
    print(f"  출력 경로: output/daily/{date_str}/")

    total_tasks = 0

    if not skip_macro:
        print(f"\n  [매크로 분석]")
        print(f"    -> fed_macro_analyzer.py")
        total_tasks += 1

    for sector_key, sector_info in sectors.items():
        sector_name = sector_info["name"]
        tickers = sector_info["tickers"]
        print(f"\n  [{sector_name}]")
        for ticker, name in tickers.items():
            print(f"    -> {name}({ticker}): 재무분석, 레짐분석, 심리분석")
            total_tasks += 3
        if not skip_portfolio:
            ticker_list = ",".join(tickers.keys())
            print(f"    -> 포트폴리오 분석: {ticker_list}")
            total_tasks += 1

    print(f"\n  총 실행 작업: {total_tasks}건")
    print(f"  예상 소요 시간: 약 {total_tasks * 20 // 60}분 ~ {total_tasks * 40 // 60}분")
    print("=" * 60)


# ====================================================================
# 7. GitHub Pages 자동 업데이트
# ====================================================================
def _auto_push_dashboard(date_str):
    """docs/data.json을 자동 커밋하고 GitHub에 push합니다."""
    try:
        # git이 초기화되어 있는지 확인
        git_dir = os.path.join(BASE_DIR, ".git")
        if not os.path.isdir(git_dir):
            return

        # remote가 설정되어 있는지 확인
        check = subprocess.run(
            ["git", "remote"],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=10
        )
        if not check.stdout.strip():
            return

        # docs/data.json 스테이징 및 커밋
        subprocess.run(
            ["git", "add", "docs/data.json"],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=10
        )

        # 변경사항이 있는지 확인
        diff_check = subprocess.run(
            ["git", "diff", "--cached", "--quiet"],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=10
        )
        if diff_check.returncode == 0:
            # 변경사항 없음
            return

        formatted = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
        subprocess.run(
            ["git", "commit", "-m", f"Update dashboard data ({formatted})"],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=30
        )

        push_result = subprocess.run(
            ["git", "push"],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=60
        )
        if push_result.returncode == 0:
            print("  [GitHub Pages] 대시보드 업데이트 push 완료")
        else:
            print(f"  [GitHub Pages] push 실패: {push_result.stderr.strip()}")

    except Exception:
        print("  [GitHub Pages] 자동 push 중 오류 발생 (수동 push 필요)")


# ====================================================================
# 8. 종합 요약 생성
# ====================================================================
def generate_summary(date_str, macro_info, sector_results, portfolio_results, daily_dir):
    """종합요약 텍스트 파일 생성."""
    filepath = os.path.join(daily_dir, f"종합요약_{date_str}.txt")

    lines = []
    lines.append(f"=== 일일 섹터 분석 요약 ({date_str[:4]}-{date_str[4:6]}-{date_str[6:]}) ===")
    lines.append("")

    # 매크로 환경
    lines.append("[매크로 환경]")
    if macro_info:
        lines.append(f"- 경기 사이클: {macro_info.get('cycle_stage', 'N/A')}")
        lines.append(f"- 금리: {macro_info.get('rate_cycle', 'N/A')} (FFR {macro_info.get('ffr', 'N/A')})")
        lines.append(f"- 인플레이션: {macro_info.get('inflation_status', 'N/A')} (CPI {macro_info.get('cpi', 'N/A')})")
    else:
        lines.append("- 매크로 분석 건너뜀 또는 실패")
    lines.append("")

    # 섹터별 요약
    lines.append("[섹터별 요약]")
    for sector_key, results in sector_results.items():
        sector_name = results.get("name", sector_key)
        lines.append(f"{sector_name}:")
        ticker_results = results.get("tickers", {})
        for ticker, info in ticker_results.items():
            name = info.get("name", ticker)
            regime = info.get("regime", "N/A")
            sentiment_label = info.get("sentiment_label", "N/A")
            sentiment_score = info.get("sentiment_score", "N/A")
            if sentiment_score != "N/A":
                lines.append(f"  {name} - 레짐: {regime} / 심리: {sentiment_label}({sentiment_score})")
            else:
                lines.append(f"  {name} - 레짐: {regime} / 심리: {sentiment_label}")
        lines.append("")

    # 포트폴리오 하이라이트
    if portfolio_results:
        lines.append("[포트폴리오 하이라이트]")
        best_sharpe = None
        best_sector = None
        worst_sharpe = None
        worst_sector = None

        for sector_key, info in portfolio_results.items():
            sharpe_str = info.get("sharpe", "N/A")
            if sharpe_str == "N/A":
                continue
            try:
                sharpe_val = float(sharpe_str)
            except (ValueError, TypeError):
                continue
            if best_sharpe is None or sharpe_val > best_sharpe:
                best_sharpe = sharpe_val
                best_sector = info.get("name", sector_key)
            if worst_sharpe is None or sharpe_val < worst_sharpe:
                worst_sharpe = sharpe_val
                worst_sector = info.get("name", sector_key)

        if best_sector:
            lines.append(f"- 최고 Sharpe 섹터: {best_sector} ({best_sharpe:.2f})")
        if worst_sector:
            lines.append(f"- 최저 Sharpe 섹터: {worst_sector} ({worst_sharpe:.2f})")
        if not best_sector:
            lines.append("- Sharpe 비율 데이터 없음")
        lines.append("")

    # 생성 시각
    lines.append(f"[생성 시각] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))

    return filepath


# ====================================================================
# 8. 메인 실행
# ====================================================================
def main():
    parser = argparse.ArgumentParser(
        description="일일 섹터별 종합 분석 시스템",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument('--sectors', default=None,
                        help='분석할 섹터 (쉼표 구분, 예: IT,DEFENSE). 생략시 전체')
    parser.add_argument('--skip-macro', action='store_true',
                        help='매크로 분석 건너뛰기')
    parser.add_argument('--skip-portfolio', action='store_true',
                        help='포트폴리오 분석 건너뛰기')
    parser.add_argument('--reset-config', action='store_true',
                        help='섹터 설정 파일 초기화')
    parser.add_argument('--dry-run', action='store_true',
                        help='실행 계획만 출력')
    args = parser.parse_args()

    # 실행 시각
    start_time = time.time()
    now = datetime.now()
    date_str = now.strftime('%Y%m%d')
    date_display = now.strftime('%Y-%m-%d %H:%M:%S')

    print("\n" + "=" * 60)
    print("  일일 섹터별 종합 분석 시스템")
    print(f"  실행 시각: {date_display}")
    print("=" * 60)

    # 설정 로드
    all_sectors = ensure_config(reset=args.reset_config)
    if args.reset_config:
        print("  설정 초기화 완료.")
        if not args.dry_run and not args.sectors:
            # reset-config만 요청한 경우 종료하지 않고 계속 진행
            pass

    # 섹터 필터링
    if args.sectors:
        requested = [s.strip().upper() for s in args.sectors.split(',')]
        sectors = {}
        for key in requested:
            if key in all_sectors:
                sectors[key] = all_sectors[key]
            else:
                print(f"  [경고] 알 수 없는 섹터: {key} (건너뜀)")
                print(f"         사용 가능: {', '.join(all_sectors.keys())}")
        if not sectors:
            print("  [오류] 유효한 섹터가 없습니다.")
            sys.exit(1)
    else:
        sectors = all_sectors

    # 드라이런
    if args.dry_run:
        print_dry_run(sectors, args.skip_macro, args.skip_portfolio, date_str)
        return

    # 출력 디렉토리 생성
    daily_dir = create_output_dirs(date_str, sectors, args.skip_macro, args.skip_portfolio)
    # output 폴더도 보장
    Path(SCRIPT_OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    # 에러 로거
    error_logger = ErrorLogger()

    # 전체 진행 카운트 (매크로 1 + 섹터 수)
    total_steps = len(sectors)
    if not args.skip_macro:
        total_steps += 1
    current_step = 0
    total_files = 0

    # 결과 저장
    macro_info = None
    sector_results = {}
    portfolio_results = {}

    # ----------------------------------------------------------------
    # Step 1: 매크로 분석
    # ----------------------------------------------------------------
    if not args.skip_macro:
        current_step += 1
        print(f"\n[{current_step}/{total_steps}] 매크로 경제 분석...")

        result = run_script(SCRIPTS["macro"], [], "매크로 분석")

        if result["success"]:
            print(f"  완료 ({format_elapsed(result['elapsed'])})")
            macro_info = parse_macro_stdout(result["stdout"])

            # 파일 이동
            macro_file = FILE_PATTERNS["macro"].format(date=date_str)
            dest_dir = os.path.join(daily_dir, "macro")
            moved = move_output_file(macro_file, dest_dir)
            if moved:
                total_files += 1
        else:
            print(f"  실패 ({format_elapsed(result['elapsed'])})")
            error_logger.add("매크로", "", "fed_macro_analyzer.py",
                             "실행 실패", result["stderr"])

    # ----------------------------------------------------------------
    # Step 1.5: 지정학 리스크 분석
    # ----------------------------------------------------------------
    if not args.skip_macro:
        print(f"\n  |- 지정학 리스크 분석...")

        result = run_script(SCRIPTS["geopolitical"], [], "지정학 리스크 분석")

        if result["success"]:
            print(f"     완료 ({format_elapsed(result['elapsed'])})")
            geo_file = FILE_PATTERNS["geopolitical"].format(date=date_str)
            dest_dir = os.path.join(daily_dir, "macro")
            moved = move_output_file(geo_file, dest_dir)
            if moved:
                total_files += 1
        else:
            print(f"     실패 ({format_elapsed(result['elapsed'])})")
            error_logger.add("지정학", "", "geopolitical_analyzer.py",
                             "실행 실패", result["stderr"])

    # ----------------------------------------------------------------
    # Step 2 & 3: 섹터별 분석 + 포트폴리오
    # ----------------------------------------------------------------
    for sector_key, sector_info in sectors.items():
        current_step += 1
        sector_name = sector_info["name"]
        safe_sector_name = sector_name.replace("/", "_")
        tickers = sector_info["tickers"]

        print(f"\n[{current_step}/{total_steps}] {sector_name} 섹터 분석...")
        sector_start = time.time()

        sector_results[sector_key] = {
            "name": sector_name,
            "tickers": {},
        }

        dest_dir = os.path.join(daily_dir, safe_sector_name)

        # 각 종목 분석
        for ticker, ticker_name in tickers.items():
            sector_results[sector_key]["tickers"][ticker] = {
                "name": ticker_name,
                "regime": "N/A",
                "sentiment_label": "N/A",
                "sentiment_score": "N/A",
            }

            # (a) 재무 분석
            print(f"  |- {ticker_name}({ticker}) 재무분석...", end=" ", flush=True)
            result = run_script(SCRIPTS["financial"], ["--ticker", ticker],
                                f"{ticker_name} 재무분석")
            if result["success"]:
                print("완료")
                # 파일 이동 (금융분석 xlsx + metrics json)
                fname = FILE_PATTERNS["financial"].format(
                    ticker=ticker.upper() if not ticker[0].isdigit() else ticker,
                    date=date_str,
                )
                moved = move_output_file(fname, dest_dir)
                if moved:
                    total_files += 1
                # metrics json도 이동
                json_fname = FILE_PATTERNS["metrics"].format(
                    ticker=ticker.upper() if not ticker[0].isdigit() else ticker,
                )
                moved_json = move_output_file(json_fname, dest_dir)
                if moved_json:
                    total_files += 1
            else:
                print("실패")
                error_logger.add(sector_name, ticker, "financial_analyzer.py",
                                 f"{ticker_name}({ticker}) 재무분석 실패", result["stderr"])

            # (b) HMM 레짐 분석
            print(f"  |- {ticker_name}({ticker}) 레짐분석...", end=" ", flush=True)
            result = run_script(SCRIPTS["hmm"], ["--ticker", ticker],
                                f"{ticker_name} 레짐분석")
            if result["success"]:
                print("완료")
                regime = parse_regime_stdout(result["stdout"])
                sector_results[sector_key]["tickers"][ticker]["regime"] = regime
                # 파일 이동
                fname = FILE_PATTERNS["hmm"].format(
                    ticker=ticker.upper() if not ticker[0].isdigit() else ticker,
                    date=date_str,
                )
                moved = move_output_file(fname, dest_dir)
                if moved:
                    total_files += 1
            else:
                print("실패")
                error_logger.add(sector_name, ticker, "hmm_regime_detector.py",
                                 f"{ticker_name}({ticker}) 레짐분석 실패", result["stderr"])

            # (c) 심리 분석
            print(f"  |- {ticker_name}({ticker}) 심리분석...", end=" ", flush=True)
            result = run_script(SCRIPTS["sentiment"], ["--ticker", ticker],
                                f"{ticker_name} 심리분석")
            if result["success"]:
                print("완료")
                label, score = parse_sentiment_stdout(result["stdout"])
                sector_results[sector_key]["tickers"][ticker]["sentiment_label"] = label
                sector_results[sector_key]["tickers"][ticker]["sentiment_score"] = score
                # 파일 이동
                fname = FILE_PATTERNS["sentiment"].format(
                    ticker=ticker.upper() if not ticker[0].isdigit() else ticker,
                    date=date_str,
                )
                moved = move_output_file(fname, dest_dir)
                if moved:
                    total_files += 1
            else:
                print("실패")
                error_logger.add(sector_name, ticker, "news_sentiment_analyzer.py",
                                 f"{ticker_name}({ticker}) 심리분석 실패", result["stderr"])

        # 섹터 포트폴리오 분석
        if not args.skip_portfolio and len(tickers) >= 2:
            ticker_list = ",".join(tickers.keys())
            print(f"  |- {sector_name} 포트폴리오 분석...", end=" ", flush=True)
            result = run_script(SCRIPTS["portfolio"], ["--tickers", ticker_list],
                                f"{sector_name} 포트폴리오")
            if result["success"]:
                print("완료")
                sharpe = parse_portfolio_stdout(result["stdout"])
                portfolio_results[sector_key] = {
                    "name": sector_name,
                    "sharpe": sharpe,
                }
                # 파일 이동
                port_fname = FILE_PATTERNS["portfolio"].format(date=date_str)
                port_dest_dir = os.path.join(daily_dir, "포트폴리오")
                # 포트폴리오 파일은 섹터별로 이름 변경
                src_path = os.path.join(SCRIPT_OUTPUT_DIR, port_fname)
                if os.path.exists(src_path):
                    dest_fname = f"{safe_sector_name}_포트폴리오_{date_str}.xlsx"
                    dst_path = os.path.join(port_dest_dir, dest_fname)
                    if os.path.exists(dst_path):
                        os.remove(dst_path)
                    shutil.move(src_path, dst_path)
                    total_files += 1
            else:
                print("실패")
                error_logger.add(sector_name, "", "portfolio_analyzer.py",
                                 f"{sector_name} 포트폴리오 분석 실패", result["stderr"])

        sector_elapsed = time.time() - sector_start
        print(f"  섹터 완료 ({format_elapsed(sector_elapsed)})")

    # ----------------------------------------------------------------
    # Step 4: 종합 요약 생성
    # ----------------------------------------------------------------
    print(f"\n종합 요약 생성 중...", end=" ", flush=True)
    summary_path = generate_summary(date_str, macro_info, sector_results,
                                    portfolio_results, daily_dir)
    total_files += 1
    print("완료")

    # Excel 종합보고서 생성
    summary_excel_script = os.path.join(BASE_DIR, "generate_summary_excel.py")
    if os.path.exists(summary_excel_script):
        print(f"\nExcel 종합보고서 생성 중...", end=" ", flush=True)
        excel_result = run_script(
            summary_excel_script,
            ["--date", date_str, "--daily-dir", daily_dir],
            "Excel 종합보고서",
        )
        if excel_result and excel_result.returncode == 0:
            total_files += 1
            print("완료")
        else:
            print("실패 (Excel 보고서 생략)")
    else:
        print(f"\n[경고] generate_summary_excel.py 파일을 찾을 수 없습니다: {summary_excel_script}")

    # 대시보드 데이터 생성 (docs/data.json)
    dashboard_script = os.path.join(BASE_DIR, "generate_dashboard_data.py")
    if os.path.exists(dashboard_script):
        print(f"\n대시보드 데이터 생성 중...", end=" ", flush=True)
        dash_result = run_script(
            dashboard_script,
            ["--date", date_str, "--daily-dir", daily_dir],
            "대시보드 데이터",
        )
        if dash_result and dash_result.returncode == 0:
            total_files += 1
            print("완료")
            # GitHub Pages 자동 업데이트 (push)
            _auto_push_dashboard(date_str)
        else:
            print("실패 (대시보드 데이터 생략)")
    else:
        print(f"\n[경고] generate_dashboard_data.py 파일을 찾을 수 없습니다.")

    # 에러 로그 저장
    error_log_path = os.path.join(daily_dir, "errors.log")
    error_logger.save(error_log_path)
    if error_logger.errors:
        total_files += 1

    # ----------------------------------------------------------------
    # 최종 보고
    # ----------------------------------------------------------------
    total_elapsed = time.time() - start_time

    print("\n" + "=" * 60)
    print("  분석 완료!")
    print(f"  총 소요 시간: {format_elapsed(total_elapsed)}")
    print(f"  생성 파일: {total_files}개")
    print(f"  저장 위치: output/daily/{date_str}/")
    if error_logger.errors:
        print(f"  오류: {len(error_logger.errors)}건 (errors.log 참조)")
    print("=" * 60)

    error_logger.print_summary()


if __name__ == "__main__":
    main()
