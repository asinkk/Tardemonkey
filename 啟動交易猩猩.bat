@echo off
chcp 65001 >nul
title Tardemonkey local
cd /d "%~dp0"
where py >nul 2>nul && (py -3 local\server.py %*) || (python local\server.py %*)
echo.
echo Server stopped. Press any key to close.
pause >nul
