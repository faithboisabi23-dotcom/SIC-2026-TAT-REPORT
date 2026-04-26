@echo off
echo ============================================
echo  TAT Dashboard - Data Refresh
echo ============================================
echo.

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.9 or higher.
    pause
    exit /b 1
)

:: Install / upgrade dependencies silently
echo Installing dependencies...
pip install -r "%~dp0requirements.txt" --quiet

echo.
echo Running data export...
python "%~dp0scripts\export_dashboard_json.py"

if errorlevel 1 (
    echo.
    echo ERROR: Script failed. Check the output above for details.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Done! dashboard/data/ has been updated.
echo  Commit and push to refresh GitHub Pages.
echo ============================================
pause