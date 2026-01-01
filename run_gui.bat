@echo off
title Experiment Analyzer GUI
echo ========================================
echo   Experiment Analyzer GUI
echo ========================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python and try again.
    echo.
    pause
    exit /b 1
)

echo Starting GUI...
echo.

REM Change to the script directory
cd /d "%~dp0"

REM Run the GUI
python experiment_analyzer_gui.py

REM If there's an error, keep the window open
if errorlevel 1 (
    echo.
    echo ========================================
    echo An error occurred while running the GUI.
    echo ========================================
    echo.
    pause
)

