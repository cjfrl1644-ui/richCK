@echo off
REM ===============================================================
REM  ASCII ONLY - no Korean anywhere in this file, not even in
REM  comments or filenames. cmd.exe reads .bat in the system
REM  codepage (CP949), not UTF-8, so any non-ASCII byte shifts the
REM  parser and every following line fails.
REM  Korean guidance lives in the guide text file and in run.py.
REM  Run this ONCE on a new computer.
REM ===============================================================
chcp 65001 > nul
cd /d "%~dp0"

echo.
echo  === SETUP - run this once on a new PC ===
echo.

python --version
if errorlevel 1 (
  echo.
  echo  [!] Python is not installed.
  echo.
  echo      1. Open  https://www.python.org/downloads/
  echo      2. Download and install Python
  echo      3. IMPORTANT: tick "Add python.exe to PATH" during setup
  echo      4. Run this file again
  echo.
  pause
  exit /b 1
)

echo.
echo  Installing required libraries. Please wait...
echo.
python -m pip install --quiet --upgrade openpyxl cryptography
if errorlevel 1 (
  echo.
  echo  [!] Install failed. Check your internet connection.
  pause
  exit /b 1
)

echo.
echo  Done. Now double-click the RUN bat file in this folder.
echo.
pause
