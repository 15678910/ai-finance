"""KRX 로그인 사전 점검 — 수집기를 돌리기 전에 로그인을 딱 한 번만 시도한다.

Context: KRX 수집기는 종목마다 로그인을 두 번씩 한다(pykrx import 시 1번 + 본문 1번).
  비밀번호가 틀리면 실행 한 번에 실패가 약 10번 쌓이고, KRX는 실패가 한도를 넘으면
  아이디/비밀번호 로그인을 잠근다(2026-11-16 실제로 잠김).
Solution: update_kospi_valuation.bat이 이 점검을 먼저 돌리고, 실패하면 그 실행 동안
  KRX_ID/KRX_PW를 비운다. 수집기는 자격증명이 없으면 로그인 없이 건너뛴다.

pykrx를 import하지 않는다(import만으로 로그인을 한 번 더 시도하기 때문).
요청 형식은 pykrx/website/comm/auth.py의 login_krx와 같다.
비밀번호와 아이디는 출력하지 않는다.

종료 코드: 0 정상(CD001) · 2 서버가 거부 · 3 네트워크/응답 오류 · 4 자격증명 미설정
수동 진단: C:\\Python313\\python.exe krx_login_check.py --registry
  (--registry: 이미 열린 창의 예전 값 대신 레지스트리에 저장된 최신 사용자 환경변수를 쓴다)
"""
import os
import sys

import requests

PAGE = "https://data.krx.co.kr/contents/MDC/COMS/client/MDCCOMS001.cmd"
JSP = "https://data.krx.co.kr/contents/MDC/COMS/client/view/login.jsp?site=mdc"
URL = "https://data.krx.co.kr/contents/MDC/COMS/client/MDCCOMS001D1.cmd"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")
CODES = {"CD001": "정상", "CD010": "비밀번호 변경 필요", "CD011": "중복 로그인"}


def _creds(from_registry):
    if from_registry and os.name == "nt":
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as k:
            def q(n):
                try:
                    return winreg.QueryValueEx(k, n)[0]
                except FileNotFoundError:
                    return ""
            return q("KRX_ID"), q("KRX_PW")
    return os.environ.get("KRX_ID", ""), os.environ.get("KRX_PW", "")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    login_id, login_pw = _creds("--registry" in sys.argv)
    if not (login_id and login_pw):
        print("[KRX 점검] KRX_ID 또는 KRX_PW 미설정")
        return 4
    try:
        s = requests.Session()
        s.get(PAGE, headers={"User-Agent": UA}, timeout=15)
        s.get(JSP, headers={"User-Agent": UA, "Referer": PAGE}, timeout=15)
        payload = {"mbrNm": "", "telNo": "", "di": "", "certType": "",
                   "mbrId": login_id, "pw": login_pw}
        hdr = {"User-Agent": UA, "Referer": PAGE}
        d = s.post(URL, data=payload, headers=hdr, timeout=15).json()
        if d.get("_error_code") == "CD011":          # 중복 로그인 — pykrx와 같이 skipDup로 한 번 더
            payload["skipDup"] = "Y"
            d = s.post(URL, data=payload, headers=hdr, timeout=15).json()
    except (requests.RequestException, ValueError) as e:
        print(f"[KRX 점검] 네트워크/응답 오류: {type(e).__name__}")
        return 3
    code = d.get("_error_code") or "?"
    msg = str(d.get("_error_message") or "")[:60]   # 로그에 남으므로 짧게(서버 문구에 계정 정보가 섞일 수 있음)
    print(f"[KRX 점검] {code} ({CODES.get(code, '기타 거부')}) {msg}".rstrip())
    return 0 if code == "CD001" else 2


if __name__ == "__main__":
    sys.exit(main())
