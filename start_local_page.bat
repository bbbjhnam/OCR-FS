@echo off
cd /d "%~dp0"
set "PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PY%" set "PY=python"
start "Financial PDF Local Server" "%PY%" -m http.server 8765 --bind 127.0.0.1
timeout /t 1 >nul
start "" "http://127.0.0.1:8765/index.html"
