@echo off
REM Windows 工作排程器用：每天執行一次，產生 docs\index.html
cd /d %~dp0
python txo_report.py
