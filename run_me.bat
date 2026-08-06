@echo off
:: ============================================================
::  Abhinav Group — Email Signature Generator
::  Double-click this file to generate all employee signatures
::  No technical knowledge required
:: ============================================================
title Abhinav Group — Signature Generator
color 0B
cls

echo.
echo  ============================================================
echo    Abhinav Group -- Email Signature Generator
echo  ============================================================
echo.
echo  This tool will:
echo    1. Read employee data from employees\employees.xlsx
echo    2. Generate one HTML signature per employee
echo    3. Open a preview so you can check all signatures
echo.
echo  ============================================================
echo.

:: ── Step 1: Check Python ─────────────────────────────────────
echo  [Checking] Looking for Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Python is NOT installed on this computer.
    echo.
    echo  To fix this:
    echo    1. Go to: https://www.python.org/downloads/
    echo    2. Download and install Python (click "Add to PATH" during install)
    echo    3. Restart your computer
    echo    4. Double-click this file again
    echo.
    echo  Press any key to exit...
    pause >nul
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo  [OK] %%v found
echo.

:: ── Step 2: Install / update required packages ───────────────
echo  [Installing] Checking required packages (openpyxl, Pillow)...
python -m pip install openpyxl Pillow --quiet --disable-pip-version-check 2>nul
if %errorlevel% neq 0 (
    echo  [WARNING] Could not install packages automatically.
    echo            Trying to continue anyway...
)
echo  [OK] Packages ready
echo.

:: ── Step 3: Run the generator ────────────────────────────────
echo  [Generating] Creating signatures from Excel data...
echo  ------------------------------------------------------------
python generate_signatures.py
if %errorlevel% neq 0 (
    echo.
    echo  ============================================================
    echo  [ERROR] Something went wrong. Read the message above.
    echo.
    echo  Common fixes:
    echo    - Make sure employees\employees.xlsx is closed in Excel
    echo    - Make sure images\logo.jpg exists
    echo    - Make sure all icon files exist in icons\ folder
    echo  ============================================================
    echo.
    echo  Press any key to exit...
    pause >nul
    exit /b 1
)

:: ── Step 4: Open preview ─────────────────────────────────────
echo  [Opening] Launching preview in your browser...
start "" "preview_all.html"

echo.
echo  ============================================================
echo   DONE! Your browser should show all signatures now.
echo.
echo   To update a signature:
echo     1. Open  employees\employees.xlsx
echo     2. Change the details
echo     3. Save the file
echo     4. Double-click this file again
echo.
echo   To add a new employee:
echo     1. Open  employees\employees.xlsx
echo     2. Add a new row at the bottom
echo     3. Save the file
echo     4. Double-click this file again
echo.
echo   Individual signature files are in the  output\  folder.
echo  ============================================================
echo.
echo  Press any key to close this window...
pause >nul
