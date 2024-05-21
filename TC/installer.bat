@echo off
REM Check if Python is installed
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python is not installed. Please install Python and try again.
    exit /b 1
)

REM Create the Timecard directory if it doesn't exist
if not exist "Timecard" (
    mkdir Timecard
)

REM Install the required packages to the Timecard directory
python -m pip install --upgrade pip
python -m pip install --target=Timecard -r requirements.txt

echo.
echo Dependencies installed successfully in the Timecard directory.
echo.

echo Setup complete.
pause
