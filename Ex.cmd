@echo off
chcp 65001 >nul

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found on this computer.
    echo Please install it from: https://www.python.org/downloads/
    echo During setup, check the box "Add Python to PATH" before clicking Install.
    echo After installing, double-click this file again.
    echo.
    pause
    exit /b 1
)

python "%~dp0pdf_to_excel.py" %*
 