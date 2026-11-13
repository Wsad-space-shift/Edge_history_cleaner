@echo off
setlocal
title Edge Site History Cleaner
set "SCRIPT=%~dp0edge_history_cleaner.py"

where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw "%SCRIPT%" %*
    exit /b 0
)
where pyw >nul 2>&1
if %errorlevel%==0 (
    start "" pyw -3 "%SCRIPT%" %*
    exit /b 0
)
where python >nul 2>&1
if %errorlevel%==0 (
    python "%SCRIPT%" %*
    exit /b %errorlevel%
)

echo [X] Python 3 was not found.
echo     Install it from https://www.python.org/downloads/
echo     and tick "Add python.exe to PATH" during setup.
pause
exit /b 1