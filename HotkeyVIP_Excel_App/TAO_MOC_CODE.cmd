@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Tao moc phien ban HotkeyVIP Excel App

set /p MESSAGE=Mo ta thay doi: 
if "%MESSAGE%"=="" (
  echo Can nhap mo ta thay doi.
  pause
  exit /b 1
)

set /p TAG=Tag phien ban, vi du app-v1.5.2: 
if "%TAG%"=="" (
  echo Can nhap tag phien ban.
  pause
  exit /b 1
)

git add -A
git commit -m "%MESSAGE%"
if errorlevel 1 (
  echo Khong tao duoc commit. Kiem tra thong bao phia tren.
  pause
  exit /b 1
)

git tag -a "%TAG%" -m "%MESSAGE%"
if errorlevel 1 (
  echo Commit da tao nhung tag bi loi. Kiem tra tag co bi trung khong.
  pause
  exit /b 1
)

echo.
echo Da tao moc %TAG%.
git log -1 --oneline --decorate
pause

