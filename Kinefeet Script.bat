@echo off
:: Step 0: Change directory to the script's own directory
cd /d %~dp0

:: Step 1: Check if Python is installed
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Python is not installed. Please install Python and try again.
    exit /b 1
)

:: Step 2: Create a virtual environment (if it doesn't already exist)
if not exist env (
    python -m venv env
    echo Virtual environment 'env' created.
    call env\Scripts\activate
    pip install -r requirements.txt
    echo Requirements installed.
    mkdir data
    cls
)

:: Step 3: Activate the virtual environment
call env\Scripts\activate

:: Step 4: Run the Python script
if exist get_data.py (
    python get_data.py
) else (
    echo get_data.py not found. Exiting.
)

:: Pause to prevent window from closing immediately (optional)
pause

