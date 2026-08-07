@echo off
rem Dang nhap lan luot cho 14 profile submit.
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_he_thong\setup_submit_edge_profiles.ps1"
