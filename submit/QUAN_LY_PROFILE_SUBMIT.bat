@echo off
rem Mo giao dien quan ly anh xa 14 Edge profile, Gmail va ten mien.
cd /d "%~dp0"
start "" /b powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0_he_thong\manage_submit_edge_profiles.ps1"
exit /b
