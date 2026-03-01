@echo off
chcp 65001 >nul
echo ============================================
echo  S&P 500 Daily Update - %date% %time%
echo ============================================

set SCRIPT_DIR=%~dp0
set PYTHON=C:\Python314\python.exe
set LOG=%SCRIPT_DIR%update_log.txt

echo [%date% %time%] === Daily Update START === >> "%LOG%"

:: Step 1: 데이터 수집
echo [1/2] Fetching S&P 500 data...
"%PYTHON%" "%SCRIPT_DIR%daily_sp500_fetcher.py"
if %errorlevel% neq 0 (
    echo [%date% %time%] ERROR: Fetcher failed >> "%LOG%"
    echo ERROR: Data fetch failed!
    goto :end
)
echo [%date% %time%] Fetcher completed >> "%LOG%"

:: Step 2: 대시보드 생성
echo [2/2] Generating dashboard...
"%PYTHON%" "%SCRIPT_DIR%generate_dashboard.py"
if %errorlevel% neq 0 (
    echo [%date% %time%] ERROR: Dashboard generator failed >> "%LOG%"
    echo ERROR: Dashboard generation failed!
    goto :end
)
echo [%date% %time%] Dashboard generated >> "%LOG%"

:: Step 3: 대시보드 자동 열기
echo Opening dashboard...
start "" "%SCRIPT_DIR%dashboard_today.html"
echo [%date% %time%] Dashboard opened >> "%LOG%"

:: Step 4: Windows 알림
echo Sending notification...
powershell -Command "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null; $xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02); $texts = $xml.GetElementsByTagName('text'); $texts[0].AppendChild($xml.CreateTextNode('📊 QUANTUM PM 업데이트 완료')) > $null; $texts[1].AppendChild($xml.CreateTextNode('대시보드가 최신 데이터로 갱신되었습니다.')) > $null; $toast = [Windows.UI.Notifications.ToastNotification]::new($xml); [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('QUANTUM PM').Show($toast)" 2>nul
echo [%date% %time%] Notification sent >> "%LOG%"

echo ============================================
echo  UPDATE COMPLETE
echo ============================================
echo [%date% %time%] === Daily Update DONE === >> "%LOG%"

:end
