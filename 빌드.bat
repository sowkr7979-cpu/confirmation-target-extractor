@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
py -3 -m pip install --quiet -r requirements.txt pyinstaller || goto :fail
py -3 -m PyInstaller --noconfirm --clean build.spec || goto :fail
xcopy /E /I /Y config dist\config >/dev/null || goto :fail
xcopy /E /I /Y template dist\template >/dev/null || goto :fail
py -3 tools\build_done.py
pause
exit /b 0
:fail
echo.
echo BUILD FAILED - see messages above.
pause
exit /b 1
