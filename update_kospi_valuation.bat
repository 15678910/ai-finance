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
REM EnableDelayedExpansion: %VAR% inside an if(...) block expands at parse time,
REM before the for-loop in the same block has set it. !VAR! reads the live value.
setlocal EnableDelayedExpansion
cd /d "C:\Users\lacoi\Desktop\ai-finance"
set "PY=C:\Python313\python.exe"
set "GIT=C:\Program Files\Git\mingw64\bin\git.exe"
set "LOG=update_kospi_valuation.log"
set "PATHS=docs/kospi_valuation.json docs/short_interest.json docs/personal_buy_dist.json docs/stock_financials.json docs/stock_screener.json"

echo [%date% %time%] START >> "%LOG%"

REM 0) SAFETY — checked first, every day, even when nothing changed.
REM    An interrupted "pull --rebase" left HEAD detached on 2026-10-08. Every later
REM    run committed there, every push failed, yet this script logged "PUSHED", so
REM    the public dashboard froze for 5.5 weeks. Do NOT auto-abort a stale rebase:
REM    "rebase --abort" resets the branch to its pre-rebase point and can discard
REM    commits made since. Stop, log, alert, and leave the repair to a human.
set "STUCK="
if exist ".git\rebase-merge" set "STUCK=1"
if exist ".git\rebase-apply" set "STUCK=1"
if defined STUCK (
  echo [%date% %time%] ERROR: interrupted rebase in .git - commit/push SKIPPED. Inspect, then "git rebase --quit" or --continue >> "%LOG%"
  call :alert "KRX local update SKIPPED: interrupted git rebase in ai-finance. See update_kospi_valuation.log"
  goto end
)
set "BR="
for /f "delims=" %%b in ('"%GIT%" symbolic-ref --short -q HEAD') do set "BR=%%b"
if not defined BR (
  REM Don't just "git checkout master": commits made on the detached HEAD would be stranded.
  echo [%date% %time%] ERROR: detached HEAD - commit/push SKIPPED. First "git branch rescue-detached HEAD", then compare with master before switching >> "%LOG%"
  call :alert "KRX local update SKIPPED: ai-finance is on a detached HEAD. See update_kospi_valuation.log"
  goto end
)

REM 1) collect
"%PY%" kospi_valuation_monitor.py >> "%LOG%" 2>&1
"%PY%" short_interest_monitor.py >> "%LOG%" 2>&1
"%PY%" personal_buy_dist.py >> "%LOG%" 2>&1
"%PY%" stock_financials.py >> "%LOG%" 2>&1
"%PY%" stock_screener.py >> "%LOG%" 2>&1

REM 2) skip if nothing changed (porcelain covers modified + untracked)
set "CHANGED="
for /f %%i in ('"%GIT%" status --porcelain -- %PATHS%') do set "CHANGED=1"
if not defined CHANGED (
  echo [%date% %time%] no change, skip >> "%LOG%"
  goto end
)

REM 3) commit ONLY our five files ("-- paths" ignores anything else already staged),
REM    stash the rest of the worktree, rebase onto the remote, push, restore the stash.
"%GIT%" add -- %PATHS%
"%GIT%" commit -m "Update local KRX data (valuation/short/personal/financials) [automated]" -- %PATHS% >> "%LOG%" 2>&1
if errorlevel 1 (
  REM Without this check, rev-list below still reads 0 and the log says PUSHED with nothing committed.
  echo [%date% %time%] ERROR: commit FAILED - data NOT published >> "%LOG%"
  call :alert "KRX local update: git commit FAILED, data NOT published. See update_kospi_valuation.log"
  goto end
)

REM Pop only a stash this run created — never someone else's older stash.
set "STASH0="
set "STASH1="
for /f "delims=" %%s in ('"%GIT%" rev-parse -q --verify refs/stash 2^>nul') do set "STASH0=%%s"
"%GIT%" stash push -u -m wip-kv >nul 2>&1
for /f "delims=" %%s in ('"%GIT%" rev-parse -q --verify refs/stash 2^>nul') do set "STASH1=%%s"
set "MYSTASH="
if defined STASH1 if not "!STASH1!"=="!STASH0!" set "MYSTASH=1"

"%GIT%" pull --rebase -X theirs >> "%LOG%" 2>&1
if errorlevel 1 (
  REM This rebase started in this run, so aborting it only undoes this run's rebase.
  echo [%date% %time%] ERROR: pull --rebase failed - aborted this run's rebase, push SKIPPED >> "%LOG%"
  "%GIT%" rebase --abort >> "%LOG%" 2>&1
  call :popstash
  call :alert "KRX local update: git pull --rebase failed, data NOT published. See update_kospi_valuation.log"
  goto end
)
"%GIT%" push >> "%LOG%" 2>&1
if errorlevel 1 (
  echo [%date% %time%] ERROR: push FAILED - data NOT published >> "%LOG%"
  call :alert "KRX local update: git push FAILED, data NOT published. See update_kospi_valuation.log"
) else (
  REM Trust the ref, not the exit code: confirm nothing is left unpushed.
  set "AHEAD="
  for /f "delims=" %%a in ('"%GIT%" rev-list --count @{u}..HEAD 2^>nul') do set "AHEAD=%%a"
  if not defined AHEAD (
    echo [%date% %time%] ERROR: no upstream for !BR! - cannot confirm push >> "%LOG%"
    call :alert "KRX local update: branch has no upstream, push NOT confirmed. See update_kospi_valuation.log"
  ) else if "!AHEAD!"=="0" (
    echo [%date% %time%] PUSHED on !BR! >> "%LOG%"
  ) else (
    echo [%date% %time%] ERROR: !AHEAD! commit^(s^) still unpushed on !BR! >> "%LOG%"
    call :alert "KRX local update: commits still unpushed after git push. See update_kospi_valuation.log"
  )
)
call :popstash
goto end

:popstash
if not defined MYSTASH goto :eof
"%GIT%" stash pop >> "%LOG%" 2>&1
if not errorlevel 1 goto popdone
REM A failed pop keeps the stash entry. Leaving conflict markers in the tree would
REM break every later run (stash push and pull --rebase both refuse an unmerged tree).
REM Reset the tree to HEAD, but ONLY while our stash entry still exists: then every
REM local change is still in it and nothing is lost. Ignored files are untouched.
set "STASH2="
for /f "delims=" %%s in ('"%GIT%" rev-parse -q --verify refs/stash 2^>nul') do set "STASH2=%%s"
if "!STASH2!"=="!STASH1!" (
  "%GIT%" reset -q --hard HEAD >> "%LOG%" 2>&1
  echo [%date% %time%] WARN: stash pop conflicted - tree reset to HEAD, local changes kept in stash@{0} [wip-kv] >> "%LOG%"
  call :alert "KRX local update: data pushed, but your local edits conflicted with it. They are saved in git stash@{0} (wip-kv)."
) else (
  echo [%date% %time%] ERROR: stash pop failed and stash ref changed - tree left as is, inspect manually >> "%LOG%"
  call :alert "KRX local update: git stash pop failed, inspect ai-finance manually. See update_kospi_valuation.log"
)
:popdone
set "MYSTASH="
goto :eof

:alert
REM Best effort: a log line nobody reads is how the 5.5-week freeze went unnoticed.
"%PY%" -c "import sys; from core import send_message; send_message(sys.argv[1])" "%~1" >> "%LOG%" 2>&1
goto :eof

:end
echo [%date% %time%] END >> "%LOG%"
endlocal
