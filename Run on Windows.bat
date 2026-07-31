@echo off
REM Double-click this file on Windows to launch Project Gatherer.
setlocal

REM Move to the folder this launcher lives in (works from any drive letter).
cd /d "%~dp0"

REM Find a Python 3 interpreter. Try the launcher first, then python.
set "PY="
where py >nul 2>nul
if %ERRORLEVEL%==0 (
    py -3 -c "import sys" >nul 2>nul
    if %ERRORLEVEL%==0 set "PY=py -3"
)
if not defined PY (
    where python >nul 2>nul
    if %ERRORLEVEL%==0 set "PY=python"
)

if not defined PY (
    echo.
    echo   Python 3 was not found on this PC.
    echo.
    echo   Install it from https://www.python.org/downloads/  and make sure
    echo   "Add python.exe to PATH" is checked during installation.
    echo   Tkinter is included with the standard python.org installer.
    echo.
    pause
    exit /b 1
)

echo Starting Project Gatherer...
%PY% project_gatherer.py
set STATUS=%ERRORLEVEL%

if not "%STATUS%"=="0" (
    echo.
    echo   The app exited with an error (code %STATUS%).
    pause
)

endlocal
