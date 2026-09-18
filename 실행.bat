@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
where py >nul 2>nul && (py -3 "src\app.py" %*) || (python "src\app.py" %*)
if errorlevel 1 pause
