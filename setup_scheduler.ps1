# Windows 작업 스케줄러 자동 등록 스크립트
# 관리자 권한으로 실행: 우클릭 → PowerShell에서 실행

$taskName = "SP500_Daily_Update"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$batPath = Join-Path $scriptDir "daily_update.bat"

# 기존 작업 삭제 (있을 경우)
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "기존 작업 '$taskName' 삭제 완료" -ForegroundColor Yellow
}

# 트리거: 매일 오전 7:00 (한국시간, 미국장 마감 후)
$trigger = New-ScheduledTaskTrigger -Daily -At "07:00"

# 액션: 배치 파일 실행
$action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c `"$batPath`"" `
    -WorkingDirectory $scriptDir

# 설정: 인터넷 연결 대기, 실패시 재시도, 배터리에서도 실행
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 10) `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

# 작업 등록
Register-ScheduledTask `
    -TaskName $taskName `
    -Trigger $trigger `
    -Action $action `
    -Settings $settings `
    -Description "S&P 500 데이터 수집 및 대시보드 자동 생성 (매일 오전 7시)" `
    -RunLevel Highest

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host " 작업 스케줄러 등록 완료!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host " 작업명: $taskName"
Write-Host " 실행시간: 매일 오전 7:00"
Write-Host " 배치파일: $batPath"
Write-Host ""
Write-Host " 확인: taskschd.msc 에서 '$taskName' 검색"
Write-Host " 수동 실행: daily_update.bat 더블클릭"
Write-Host ""
