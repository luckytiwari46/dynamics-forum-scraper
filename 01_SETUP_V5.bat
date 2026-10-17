@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title SAP Learning Video Extractor V5 - SETUP
python -m venv .venv
if errorlevel 1 (echo Python was not found. Install Python 3.11+ and enable Add Python to PATH.&pause&exit /b 1)
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install playwright
".venv\Scripts\python.exe" -m playwright install chromium
if not exist browser_profile mkdir browser_profile
if not exist video_capture_output mkdir video_capture_output
echo SETUP COMPLETE.
echo Next: run 02_LOGIN_SETUP.bat, then 03_RUN_V5.bat.
pause
