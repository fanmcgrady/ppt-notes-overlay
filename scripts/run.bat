@echo off
chcp 65001 >nul
cd /d "%~dp0\.."
python notes_overlay.py
if errorlevel 1 pause
