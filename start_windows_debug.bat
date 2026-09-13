@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Ind AS 116 Lease Accounting V7.3 - Diagnostic Launcher

echo ============================================================
echo   Ind AS 116 Lease Accounting V7.3 - Diagnostic Mode
echo ============================================================
echo.
echo Application folder:
echo %~dp0
echo.
echo Python detected by Windows:
where python
echo.
python --version
echo.

echo Starting the normal launcher...
echo If an error occurs, this window will remain open.
echo.
call "%~dp0start_windows.bat"
set "RC=%ERRORLEVEL%"
echo.
echo Launcher exit code: %RC%
echo.
pause
exit /b %RC%
