@echo off
echo ============================================================
echo   GRID FORMATTER V2 DEBUG MODE
echo ============================================================
echo.

set LOG_FILE=%~dp0AI\Logs\automatic_run_%date:~-4,4%%date:~-7,2%%date:~-10,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log.txt
set LOG_FILE=%LOG_FILE: =0%

echo Log file: %LOG_FILE%
echo.

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -Command "& { try { & '%~dp0FormatGrids.ps1' 2>&1 | Tee-Object -FilePath '%LOG_FILE%' -Append } catch { Write-Host 'FATAL ERROR:' $_.Exception.Message; Write-Host $_.ScriptStackTrace; Write-Host $_.InvocationInfo.PositionMessage; exit 1 } }"

echo.
echo ============================================================
echo   COMPLETE - Check log file: %LOG_FILE%
echo ============================================================
echo Press any key to close...
pause >nul