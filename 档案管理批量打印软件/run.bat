@echo off
cd /d %~dp0
if not exist .venv\Scripts\python.exe (
  echo .venv not found. Please create the environment first with:
  echo   python -m venv .venv
  exit /b 1
)
.venv\Scripts\python.exe -m excel2word %*
