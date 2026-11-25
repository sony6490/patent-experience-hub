@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
title PatentKB - Start

set "PORT=8766"
set "URL=http://127.0.0.1:%PORT%"
set "APPDIR=%~dp0"
set "NOBROWSER="
if /i "%~1"=="/nobrowser" set "NOBROWSER=1"

echo ==============================================
echo   PatentKB  -  Patent Experience Search
echo   Folder : "%APPDIR%"
echo   URL    : %URL%
echo ==============================================
echo.

rem ---------- 1. find a python that has the required libraries ----------
rem "python" on PATH may be a version-manager shim (a .bat): always use "call",
rem and never trust the first hit - check the libraries instead.
set "PYRUN="
if defined KBPY call :accept "%KBPY%"
if not defined PYRUN if exist "%~dp0.venv\Scripts\python.exe" call :accept "%~dp0.venv\Scripts\python.exe"
if not defined PYRUN call :try python
if not defined PYRUN call :try py -3
if not defined PYRUN call :try py
if not defined PYRUN call :try py -3.13
if not defined PYRUN call :try python3
if not defined PYRUN (
  echo [ERROR] No Python found that can import: jieba python-docx openpyxl PyMuPDF
  echo         Install the requirements first:
  echo             python -m pip install -r requirements.txt
  echo         Or set KBPY to the interpreter you want, then run this again.
  echo.
  pause
  exit /b 1
)
echo Using interpreter : %PYRUN%

rem ---------- 2. already running? ----------
call :health
if not errorlevel 1 (
  echo [OK] Service is already running - nothing to do.
  if not defined NOBROWSER start "" "%URL%"
  echo.
  ping -n 3 127.0.0.1 >nul
  exit /b 0
)

rem ---------- 3. start the web service (minimized window) ----------
echo [1/3] Starting server ...
start "PatentKB Server" /min %PYRUN% "%APPDIR%server.py"

rem ---------- 4. start the sync loop if it is not running ----------
tasklist /fi "windowtitle eq PatentKB Sync" 2>nul | find /i "python" >nul
if errorlevel 1 (
  echo [2/3] Starting sync loop ...
  start "PatentKB Sync" /min %PYRUN% "%APPDIR%sync.py" loop
) else (
  echo [2/3] Sync loop already running.
)

rem ---------- 5. wait until the API answers ----------
echo [3/3] Waiting for server to answer ...
set /a TRIES=0
:wait
call :health
if not errorlevel 1 goto ready
set /a TRIES+=1
if %TRIES% geq 60 goto failed
ping -n 2 127.0.0.1 >nul
goto wait

:ready
echo.
echo [OK] Server is ready: %URL%
if not defined NOBROWSER start "" "%URL%"
echo.
echo Keep the minimized windows "PatentKB Server" and "PatentKB Sync" open.
echo To stop the service, run the stop bat file in this folder.
ping -n 4 127.0.0.1 >nul
exit /b 0

:failed
echo.
echo [ERROR] Server did not respond within 60 seconds.
echo         Check "%APPDIR%data\server_err.log" for the traceback.
echo.
pause
exit /b 1

rem ---------- helpers ----------
:try
rem args = interpreter command (python / py -3 ...); keep it only if libraries import
set "TX="
for /f "delims=" %%I in ('call %* -c "import sys;print(sys.executable)" 2^>nul') do set "TX=%%I"
if not defined TX exit /b 0
call :accept "%TX%"
exit /b 0

:accept
rem %1 = quoted path of a python executable
if not exist %1 exit /b 0
call %1 -c "import jieba, docx, openpyxl, fitz" >nul 2>nul
if errorlevel 1 exit /b 0
set "PYRUN=%1"
exit /b 0

:health
powershell -NoProfile -Command "try{(Invoke-WebRequest '%URL%/api/stats' -UseBasicParsing -TimeoutSec 2)|Out-Null;exit 0}catch{exit 1}" >nul 2>nul
if errorlevel 1 (exit /b 1) else (exit /b 0)
