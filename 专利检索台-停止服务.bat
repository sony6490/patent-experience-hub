@echo off
setlocal EnableExtensions
chcp 65001 >nul
title PatentKB - Stop
set "KBAPP=%~dp0"
echo Stopping services started from "%KBAPP%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=$env:KBAPP.TrimEnd('\'); $port=8766; try{ $cfg=(Get-Content (Join-Path $d 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json); if($cfg.server_port){$port=[int]$cfg.server_port} }catch{}; foreach($n in @('server','sync')){ $f=Join-Path $d ('data\'+$n+'.pid'); if(Test-Path $f){ $id=(Get-Content $f -ErrorAction SilentlyContinue | Select-Object -First 1); if($id -match '^\d+' -and (Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue)){ Write-Host ('  stopped '+$n+' PID '+$id); Stop-Process -Id ([int]$id) -Force }; Remove-Item $f -Force -ErrorAction SilentlyContinue } }; try{ $c=Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1; if($c){ $p=Get-CimInstance Win32_Process -Filter \"ProcessId=$($c.OwningProcess)\" -ErrorAction SilentlyContinue; if($p -and $p.CommandLine -match 'server\.py'){ Write-Host ('  stopped port owner PID '+$p.ProcessId); Stop-Process -Id $p.ProcessId -Force } } }catch{}; Write-Host '  done'"
echo.
ping -n 3 127.0.0.1 >nul
