@echo off
cd /d "%~dp0"
set "PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PY%" set "PY=python"
start "Local OCR Financial Server" "%PY%" "%~dp0local_ocr_server.py" 8788
timeout /t 1 >nul
start "" "http://127.0.0.1:8788/server"
