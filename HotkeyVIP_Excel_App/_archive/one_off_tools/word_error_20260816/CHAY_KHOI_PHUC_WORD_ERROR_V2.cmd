@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Khoi phuc Word V2 - 2 Edge tai truoc + 1 Word
set PYTHONUTF8=1
python "04_khoi_phuc_word_error_v2.py"
echo.
echo Da ket thuc. Nhan phim bat ky de dong cua so.
pause >nul
