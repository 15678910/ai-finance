@echo off
REM ============================================================
REM  KOSPI valuation local auto-updater (Korean IP only)
REM  KRX / INDEXerGO block GitHub datacenter IPs, so this runs
REM  on the local PC and pushes docs/kospi_valuation.json.
REM  - If KRX_ID / KRX_PW user env vars exist -> KRX login source
REM  - Otherwise -> INDEXerGO source (no credentials needed)
REM  Registered as Windows Scheduled Task "KOSPI Valuation Update".
REM ============================================================
setlocal
cd /d "C:\Users\lacoi\Desktop\ai-finance"
set "PY=C:\Python313\python.exe"
set "GIT=C:\Program Files\Git\mingw64\bin\git.exe"
set "LOG=update_kospi_valuation.log"

echo [%date% %time%] START >> "%LOG%"

REM 1) collect
"%PY%" kospi_valuation_monitor.py >> "%LOG%" 2>&1

REM 2) skip if unchanged
"%GIT%" diff --quiet -- docs/kospi_valuation.json
if %errorlevel%==0 (
  echo [%date% %time%] no change, skip >> "%LOG%"
  goto end
)

REM 3) commit + safe push (stash other changes, rebase-theirs, push)
"%GIT%" add docs/kospi_valuation.json
"%GIT%" commit -m "Update KOSPI valuation (local scheduler) [automated]" >> "%LOG%" 2>&1
"%GIT%" stash push -u -m wip-kv >nul 2>&1
"%GIT%" -c merge.ours.driver=true pull --rebase -X theirs >> "%LOG%" 2>&1
"%GIT%" push >> "%LOG%" 2>&1
"%GIT%" stash drop >nul 2>&1
echo [%date% %time%] PUSHED >> "%LOG%"

:end
echo [%date% %time%] END >> "%LOG%"
endlocal
