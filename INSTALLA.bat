@echo off
title Install WorkMate AI
cd /d "%~dp0"

where py >nul 2>&1
if %errorlevel%==0 (
    py -3 -m pip install --upgrade pip
    py -3 -m pip install -r requirements.txt
) else (
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
)

echo.
echo Installation completed.
pause
