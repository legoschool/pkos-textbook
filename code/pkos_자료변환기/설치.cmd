@echo off
rem  ---------------------------------------------------------------
rem   Setup for pkos_jaryo converter.  ASCII ONLY -- do not add
rem   Korean text to this file.  A .cmd that mixes Korean text with a
rem   console code page it was not saved in (949 vs 65001) desyncs the
rem   cmd.exe parser and it starts executing message lines as commands.
rem   All Korean output is printed by pkos.py, which handles UTF-8.
rem  ---------------------------------------------------------------
setlocal
cd /d "%~dp0"
title pkos - setup

echo.
echo   Looking for Python...
echo.

python -c "import sys;sys.exit(0 if sys.version_info>=(3,9) else 1)" 2>nul
if not errorlevel 1 goto :use_python
py -3 -c "import sys;sys.exit(0 if sys.version_info>=(3,9) else 1)" 2>nul
if not errorlevel 1 goto :use_py
goto :no_python

:use_python
set PY=python
goto :install

:use_py
set PY=py -3
goto :install

:install
%PY% --version
echo.
%PY% pkos.py --setup
goto :end

:no_python
echo   [!] Python 3.9+ not found.
echo.
set ans=n
set /p ans=  Install Python now? (y/n) 
if /i not "%ans%"=="y" goto :manual
echo.
echo   Installing Python. This takes a few minutes...
echo.
winget install --id Python.Python.3.13 --source winget --scope user --silent --accept-package-agreements --accept-source-agreements
if errorlevel 1 goto :manual
echo.
echo   [OK] Python installed.
echo.
echo   *** Close this window, then run this file again.  ***
echo       A new window is needed to find Python.
goto :end

:manual
echo.
echo   Please install Python from  https://www.python.org
echo   Turn ON "Add python.exe to PATH" during setup.

:end
echo.
pause
