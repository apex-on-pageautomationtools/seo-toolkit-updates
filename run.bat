@echo off
title Google Rank Checker
cd /d "%~dp0"

echo.
echo  Google Rank Checker v2.3
echo  -------------------------
echo.

:: Check embedded Python exists
if not exist "python\python.exe" (
    echo [ERROR] python\python.exe not found.
    echo         The python\ folder may be missing from this directory.
    echo.
    pause
    exit /b 1
)

:: Check main script exists
if not exist "web_app_batch.py" (
    echo [ERROR] web_app_batch.py not found.
    echo         Make sure you run this bat from inside the tool folder.
    echo.
    pause
    exit /b 1
)

echo  Starting... browser will open automatically at http://127.0.0.1:5055
echo  Keep this window open while using the tool.
echo  Press Ctrl+C here to stop.
echo.
echo  Checking for updates...
python\python.exe -s updater.py

echo.
python\python.exe -s web_app_batch.py

echo.
echo [INFO] Tool stopped.
pause
