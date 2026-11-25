@echo off
setlocal EnableExtensions
title PatentKB - Autostart
rem Register a scheduled task that starts the sync loop at logon (current user).
schtasks /Create /TN "PatentKB-Sync" /TR "\"%~dp0start_sync.bat\"" /SC ONLOGON /F
if %errorlevel%==0 (
  echo [OK] Scheduled task "PatentKB-Sync" registered.
) else (
  echo [ERROR] Registration failed. Right-click and "Run as administrator".
)
pause
