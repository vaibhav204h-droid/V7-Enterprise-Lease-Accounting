@echo off
setlocal EnableExtensions
cd /d "%~dp0"

title Ind AS 116 Lease Accounting V7.3
set "ROOT=%~dp0"
set "VENV=%ROOT%.venv"
set "PYTHON=%VENV%\Scripts\python.exe"
set "READY=%VENV%\.v73_ready"

if not exist "%ROOT%requirements.txt" goto :fail_missing
if not exist "%ROOT%app_v5.py" goto :fail_missing

REM A copied Python venv is machine-specific on Windows. Build a clean local venv
REM the first time this package is started, then reuse it on subsequent starts.
if not exist "%PYTHON%" goto :create_venv
if not exist "%READY%" goto :create_venv

REM Verify the local interpreter is actually runnable before starting Streamlit.
"%PYTHON%" -c "import sys; print('Using Python:', sys.executable)" >nul 2>&1
if errorlevel 1 goto :create_venv

goto :start_app

:create_venv
echo.
echo ============================================================
echo   Ind AS 116 Lease Accounting Platform V7.3
echo   First-run setup: creating local Python environment...
echo ============================================================
echo.

if exist "%VENV%" (
    echo Removing old/incompatible .venv...
    rmdir /s /q "%VENV%"
)

where python >nul 2>&1
if errorlevel 1 goto :fail_python

python -c "import sys; print('Base Python:', sys.executable); print('Version:', sys.version)"
python -m venv "%VENV%"
if errorlevel 1 goto :fail_venv

set "PYTHON=%VENV%\Scripts\python.exe"

if not exist "%PYTHON%" goto :fail_venv

echo.
echo Installing required packages. This is required only on first run...
"%PYTHON%" -m pip install --upgrade pip
if errorlevel 1 goto :fail_pip
"%PYTHON%" -m pip install -r "%ROOT%requirements.txt"
if errorlevel 1 goto :fail_pip

echo.> "%READY%"

echo First-run setup completed successfully.
echo.

:start_app
echo ============================================================
echo   Starting Ind AS 116 Lease Accounting Platform V7.3
echo   Browser will open automatically.
echo   Keep this window open while using the application.
echo ============================================================
echo.

"%PYTHON%" -m streamlit run "%ROOT%app_v5.py" --server.headless false --browser.gatherUsageStats false
set "APP_EXIT=%ERRORLEVEL%"

echo.
if not "%APP_EXIT%"=="0" (
    echo ============================================================
    echo   The application stopped with exit code %APP_EXIT%.
    echo   Read the error shown above and send it to me if needed.
    echo ============================================================
    pause
) else (
    echo Application closed normally.
    pause
)
exit /b %APP_EXIT%

:fail_missing
echo ERROR: Required application files were not found in:
echo %ROOT%
pause
exit /b 10

:fail_python
echo ERROR: Python was not found on this computer.
echo Install Python 3.11+ from python.org and ensure "Add Python to PATH" is enabled.
pause
exit /b 11

:fail_venv
echo ERROR: Unable to create the local Python virtual environment.
echo Please verify that Python can run with: python --version
pause
exit /b 12

:fail_pip
echo ERROR: Python packages could not be installed.
echo Check your internet connection and the error shown above.
pause
exit /b 13
