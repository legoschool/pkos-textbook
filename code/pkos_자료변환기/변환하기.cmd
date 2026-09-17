@echo off
rem  ---------------------------------------------------------------
rem   Drag a folder onto this file to convert it.  ASCII ONLY --
rem   see the note in the setup script for why.  pkos.py prints the UI.
rem  ---------------------------------------------------------------
setlocal
cd /d "%~dp0"
title pkos - convert

python -c "import sys" 2>nul
if not errorlevel 1 goto :use_python
py -3 -c "import sys" 2>nul
if not errorlevel 1 goto :use_py
goto :no_python

:use_python
set PY=python
goto :ready

:use_py
set PY=py -3
goto :ready

:ready
if "%~1"=="" goto :no_arg
%PY% pkos.py %1 --wizard
goto :end

:no_arg
echo.
echo   Drag a folder onto this file to convert it.
echo.
echo   Or from a command prompt:
echo     python pkos.py "FOLDER" --scan
echo.
echo   First time?  Run the setup script in this folder first.
goto :end

:no_python
echo.
echo   [!] Python not found.  Run the setup script in this folder first.

:end
echo.
pause
