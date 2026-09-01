@echo off
cd /d "%~dp0"
set "PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PY%" set "PY=python"
"%PY%" -c "import local_ocr_server as s; print('OCRmyPDF command:', s.find_ocrmypdf_command() or 'NOT FOUND'); print('Server URL after restart: http://127.0.0.1:8788/server'); print('Config file:', s.ROOT / 'ocrmypdf_path.txt')"
pause
