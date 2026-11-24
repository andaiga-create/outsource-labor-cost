@echo off
setlocal
cd /d "%~dp0"
echo ==========================================
echo  Outsource labor cost dashboard update
echo ==========================================

rem --- 1. find a python that has xlrd + openpyxl (several pythons may be installed) ---
set "PY="
for /f "delims=" %%P in ('where python 2^>nul') do (
  if not defined PY ("%%P" -c "import xlrd, openpyxl" >nul 2>&1 && set "PY=%%P")
)
if not defined PY (
  py -3 -c "import xlrd, openpyxl" >nul 2>&1 && set "PY=py -3"
)
if not defined PY (
  echo [!] xlrd/openpyxl not found. Installing...
  for /f "delims=" %%P in ('where python 2^>nul') do (
    if not defined PY ("%%P" -m pip install xlrd openpyxl && "%%P" -c "import xlrd, openpyxl" >nul 2>&1 && set "PY=%%P")
  )
)
if not defined PY (
  echo [ERROR] No usable Python. Install Python 3 then run: pip install xlrd openpyxl
  goto :fail
)
rem --- 0. first run: connect this folder to GitHub ---
if not exist ".git" (
  echo [0/3] First run - connecting this folder to GitHub
  git init -b main
  git remote add origin https://github.com/andaiga-create/outsource-labor-cost.git
  git fetch origin main
  git reset origin/main
)
echo [1/3] Converting xls to docs\data.json + excel report  (python: %PY%)
%PY% -W ignore scripts\build_data.py
if errorlevel 1 (
  echo [ERROR] Conversion failed. Check the xls files in AS / gimhae folders.
  goto :fail
)
if /i "%~1"=="build" (
  echo Build only. Skipped git.
  goto :end
)

rem --- 2. commit ---
echo [2/3] git commit
git add .gitignore .gitattributes README.md docs scripts update.bat
git commit -m "data update %date%"
if errorlevel 1 echo (nothing new to commit - continuing)

rem --- 3. push ---
echo [3/3] git push
git push -u origin main
if errorlevel 1 (
  echo Remote has newer commits - pulling and retrying...
  git pull --rebase origin main
  git push -u origin main
)
if errorlevel 1 (
  echo [ERROR] Push failed. Check internet / GitHub login.
  goto :fail
)
echo.
echo Done. The site updates in 1-2 minutes (Ctrl+F5 to refresh).
goto :end

:fail
echo.
echo *** FAILED - see messages above ***
:end
if /i not "%~1"=="build" pause
endlocal
