# GitHub Actions의 cron은 부하 때문에 하루 호출의 약 28%만 실행된다(2026-11-17 실측).
# 로컬 PC가 gh로 workflow_dispatch를 직접 호출해 헤더 "갱신" 시각이 늦어지지 않게 한다.
# 작업 스케줄러 "Actions Kick"이 30분마다 실행한다. 비밀은 저장하지 않는다(gh 로그인 사용).
#   :00대 -> sao-aggregator (data_refreshed_at 갱신원)
#   :30대 -> overseas-market (KOSPI 장중 KST 09~15시는 건너뜀)
#   평일 KST 18:40대 -> kr-market-gate
$ErrorActionPreference = 'Continue'
$repo = '15678910/ai-finance'
$gh   = 'C:\Users\lacoi\gh\bin\gh.exe'
$log  = Join-Path $PSScriptRoot 'kick_actions.log'
$now  = Get-Date
$kick = @()

if ($now.Minute -lt 30) {
    $kick += 'sao-aggregator.yml'
} else {
    if ($now.Hour -lt 9 -or $now.Hour -ge 16) { $kick += 'overseas-market.yml' }
}
# 평일 18:30~18:59 구간(= :30대 슬롯)에서 장 마감 게이트
if ($now.DayOfWeek -in 1..5 -and $now.Hour -eq 18 -and $now.Minute -ge 30) { $kick += 'kr-market-gate.yml' }

foreach ($wf in $kick) {
    $out = & $gh workflow run $wf --repo $repo 2>&1 | Out-String
    $ok  = ($LASTEXITCODE -eq 0)
    Add-Content -Path $log -Value ("[{0}] {1} {2} {3}" -f $now.ToString('yyyy-MM-dd HH:mm:ss'), $wf, $(if ($ok) { 'OK' } else { 'FAIL' }), $out.Trim())
}
