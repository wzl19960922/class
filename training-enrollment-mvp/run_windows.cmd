@echo off
setlocal
cd /d %~dp0
if exist "D:\conda\Scripts\conda.exe" (
  set "CONDA_EXE=D:\conda\Scripts\conda.exe"
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_windows.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] 脚本执行失败，请检查上面的输出。
  exit /b 1
)
endlocal
