@echo off
title Stop SEO Toolkit Pro
echo.
echo  Stopping SEO Toolkit Pro...
echo.

set "FOUND="
for /L %%i in (5070,1,5089) do (
    for /f "tokens=5" %%P in ('netstat -ano ^| findstr "127.0.0.1:%%i" ^| findstr LISTENING 2^>nul') do (
        taskkill /F /PID %%P >nul 2>&1
        if not errorlevel 1 set "FOUND=1"
    )
)

if defined FOUND (
    echo  SEO Toolkit Pro stopped.
) else (
    echo  SEO Toolkit Pro was not running.
)
echo.
ping -n 2 127.0.0.1 >nul
