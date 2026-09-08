@echo off
REM ===============================================================
REM  ASCII ONLY - no Korean anywhere in this file, not even in
REM  comments or filenames. cmd.exe reads .bat in the system
REM  codepage (CP949), not UTF-8, so any non-ASCII byte shifts the
REM  parser and every following line fails.
REM  All Korean messages are printed by run.py.
REM ===============================================================
chcp 65001 > nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

python -c "import openpyxl" 2>nul
if errorlevel 1 (
  echo.
  echo  [!] Not set up yet on this PC.
  echo      Double-click the SETUP bat file in this folder first.
  echo.
  pause
  exit /b 1
)

python "scripts\run.py"
if errorlevel 1 (
  echo.
  echo  [ERROR] Something went wrong. Please read the message above.
)

echo.
pause
