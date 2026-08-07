@echo off
cd /d "%~dp0"
set /p PROFILE_NUMBER=Nhap so profile can mo (1-14): 
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_he_thong\open_submit_edge_profile.ps1" -Number %PROFILE_NUMBER%
