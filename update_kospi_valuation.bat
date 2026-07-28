@echo off
REM ============================================================
REM  Local auto-updater (Korean IP only) — daily 18:30 task
REM  KRX / INDEXerGO block GitHub datacenter IPs, so this runs
REM  on the local PC and pushes the collected JSON files.
REM   1) kospi_valuation_monitor.py -> docs/kospi_valuation.json
REM   2) short_interest_monitor.py  -> docs/short_interest.json
REM      (needs KRX_ID / KRX_PW user env vars; skips if unset)
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
"%PY%" short_interest_monitor.py >> "%LOG%" 2>&1
"%PY%" personal_buy_dist.py >> "%LOG%" 2>&1

REM 2) skip if nothing changed (porcelain covers modified + untracked)
set "CHANGED="
for /f %%i in ('"%GIT%" status --porcelain -- docs/kospi_valuation.json docs/short_interest.json docs/personal_buy_dist.json') do set "CHANGED=1"
if not defined CHANGED (
  echo [%date% %time%] no change, skip >> "%LOG%"
  goto end
)

REM 3) commit + safe push (stash other changes, rebase-theirs, push)
"%GIT%" add docs/kospi_valuation.json docs/short_interest.json docs/personal_buy_dist.json
"%GIT%" commit -m "Update KOSPI valuation + short interest + personal buy dist (local scheduler) [automated]" >> "%LOG%" 2>&1
"%GIT%" stash push -u -m wip-kv >nul 2>&1
"%GIT%" pull --rebase -X theirs >> "%LOG%" 2>&1
"%GIT%" push >> "%LOG%" 2>&1
"%GIT%" stash drop >nul 2>&1
echo [%date% %time%] PUSHED >> "%LOG%"

:end
echo [%date% %time%] END >> "%LOG%"
endlocal
