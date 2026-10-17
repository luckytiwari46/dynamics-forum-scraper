@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Dynamics Communities - LOGIN SETUP

if not exist ".venv\Scripts\python.exe" (
    echo Run 01_SETUP_V5.bat first.
    pause
    exit /b 1
)

if not exist browser_profile mkdir browser_profile

".venv\Scripts\python.exe" -c "from playwright.sync_api import sync_playwright;from pathlib import Path;p=sync_playwright().start();c=p.chromium.launch_persistent_context(str(Path('browser_profile').resolve()),headless=False);pg=c.pages[0] if c.pages else c.new_page();pg.goto('https://dynamicscommunities.com/',wait_until='domcontentloaded');input('Log in to Dynamics Communities, then press ENTER here...');c.close();p.stop()"

echo LOGIN PROFILE SAVED.
pause
