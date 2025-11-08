@echo off
echo ========================================
echo Access VBA Import Automation
echo ========================================
echo.
echo Launching 32-bit PowerShell...
echo.

REM Direct execution with full path
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0Access_automaatio.ps1'"

echo.
echo ========================================
if %ERRORLEVEL% EQU 0 (
    echo Script completed successfully!
) else (
    echo Script failed with error code: %ERRORLEVEL%
)
echo ========================================
pause
