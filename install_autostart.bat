@echo off
chcp 65001 >nul
rem 开机自启同步服务（当前用户登录时启动）
schtasks /Create /TN "PatentKB-Sync" /TR "\"%~dp0start_sync.bat\"" /SC ONLOGON /F
if %errorlevel%==0 (
  echo 已注册开机自启任务 "PatentKB-Sync"
) else (
  echo 注册失败，请右键"以管理员身份运行"后重试
)
pause
