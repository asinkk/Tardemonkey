@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (py -3 local\server.py --save) || (python local\server.py --save)
echo.
pause
