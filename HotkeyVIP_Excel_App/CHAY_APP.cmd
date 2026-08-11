@echo off
setlocal
cd /d "%~dp0"

where pyw.exe >nul 2>nul
if %errorlevel%==0 (
    start "" pyw.exe -3 -m excel_audit_app.main
    exit /b 0
)

where pythonw.exe >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw.exe -m excel_audit_app.main
    exit /b 0
)

set "BUNDLED_PYTHONW=C:\Users\Admin\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
if exist "%BUNDLED_PYTHONW%" (
    start "" "%BUNDLED_PYTHONW%" -m excel_audit_app.main
    exit /b 0
)

echo Khong tim thay Python 3 de chay app.
echo Hay cai Python 3 hoac lien he nguoi tao app.
pause
