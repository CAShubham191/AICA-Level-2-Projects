@echo off
setlocal
chcp 65001 >nul
title The OG PDF Editor - Setup and Launch
cd /d "%~dp0"

echo ==========================================
echo       The OG PDF Editor - Launcher
echo ==========================================
echo.

set "OG_PYTHON="
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 (
    set "OG_PYTHON=py -3"
    goto python_found
)
python -c "import sys; assert sys.version_info.major == 3" >nul 2>&1
if not errorlevel 1 (
    set "OG_PYTHON=python"
    goto python_found
)
echo Python 3 was not found.
echo Install Python from https://www.python.org/downloads/windows/
echo During installation, select "Add Python to PATH" and include Tcl/Tk and IDLE.
echo Then run this batch file again.
goto failed

:python_found
if not exist "%~dp0The_OG_PDF_Editor_v2.py" (
    echo The_OG_PDF_Editor_v2.py was not found.
    echo Keep this batch file and the Python file in the same folder.
    goto failed
)

%OG_PYTHON% -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    echo Tkinter is missing. Modify or reinstall Python and include Tcl/Tk and IDLE.
    goto failed
)

echo Checking pip...
%OG_PYTHON% -m pip --version >nul 2>&1
if errorlevel 1 (
    %OG_PYTHON% -m ensurepip --upgrade
    if errorlevel 1 goto failed
)

echo.
echo Installing required libraries. Internet is needed for missing packages.
%OG_PYTHON% -m pip install "PyMuPDF>=1.24.0" Pillow python-docx openpyxl
if errorlevel 1 (
    echo.
    echo Dependency installation failed. Check the error above and your internet connection.
    goto failed
)

echo.
echo Starting The OG PDF Editor...
%OG_PYTHON% "%~dp0The_OG_PDF_Editor_v2.py"
if errorlevel 1 (
    echo.
    echo The editor stopped with an error. Please copy the error message above.
    goto failed
)
endlocal
exit /b 0

:failed
echo.
pause
endlocal
exit /b 1
