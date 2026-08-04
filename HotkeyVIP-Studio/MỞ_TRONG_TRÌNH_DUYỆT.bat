@echo off
cd /d "%~dp0"
powershell -NoProfile -WindowStyle Hidden -Command "if (-not (Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue)) { Start-Process pythonw.exe -ArgumentList 'app.py --no-browser' -WorkingDirectory '%~dp0' -WindowStyle Hidden }"
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765"
exit
