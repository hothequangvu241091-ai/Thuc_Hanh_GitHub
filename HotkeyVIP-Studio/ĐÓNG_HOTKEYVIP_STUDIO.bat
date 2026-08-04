@echo off
title Dong HotkeyVIP Studio
for /f "tokens=5" %%P in ('netstat -ano ^| findstr "127.0.0.1:8765" ^| findstr "LISTENING"') do (
  taskkill /PID %%P /F >nul 2>&1
)
echo Da dong HotkeyVIP Studio.
timeout /t 2 /nobreak >nul
