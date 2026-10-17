@echo off
setlocal
cd /d "%~dp0"

echo ================================================================================
echo Dynamics 365 Community - Word Exporter V5
echo ================================================================================
echo.
echo HOW IT WORKS:
echo  - You paste the archive/listing page URL.
echo  - You enter the starting and ending page numbers (or press Enter for single page).
echo  - It extracts articles found on each page in the specified range.
echo  - It generates a separate Word document for each page.
echo  - Each Word document is saved with its corresponding page number in the filename.
echo  - The browser window must stay open while extraction runs.
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    py -m venv .venv
    if errorlevel 1 goto setup_error
)

call ".venv\Scripts\activate.bat"
if not exist ".venv\Lib\site-packages\docx" (
    python -m pip install --upgrade pip
    if errorlevel 1 goto setup_error
    python -m pip install playwright python-docx
    if errorlevel 1 goto setup_error
    python -m playwright install chromium
    if errorlevel 1 goto setup_error
)

echo.
echo Starting scraper...
echo.
python "dynamics_forum_page_by_page_complete_word_V4(1).py"

echo.
echo ================================================================================
echo Scraper finished. Press any key to close.
echo ================================================================================
pause >nul
endlocal
exit /b 0

:setup_error
echo.
echo Setup failed. Check that Python is installed and available through the py command.
pause
endlocal
exit /b 1
