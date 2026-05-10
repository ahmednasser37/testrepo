@echo off
title Schedule Comparison App - Setup
color 0A

echo ============================================
echo  Schedule Comparison App - Auto Installer
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found.
    echo Please install Python from https://python.org/downloads
    echo Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)
echo [OK] Python found.

:: Create virtual environment
if not exist ".venv" (
    echo [..] Creating virtual environment...
    python -m venv .venv
)
echo [OK] Virtual environment ready.

:: Install dependencies
echo [..] Installing dependencies (this may take a minute)...
.venv\Scripts\pip install -q -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)
echo [OK] Dependencies installed.

:: Create .env file if missing
if not exist ".env" (
    echo.
    echo [..] Setting up API key...
    set /p API_KEY="Paste your OpenRouter API key and press Enter: "
    (
        echo OPENROUTER_API_KEY=%API_KEY%
        echo OPENROUTER_MODEL=deepseek/deepseek-chat-v3-0324:free
        echo SCE_CACHE_DIR=%TEMP%\sce_cache
        echo SCE_UPLOAD_MAX_MB=50
        echo PORT=5000
    ) > .env
    echo [OK] .env saved.
) else (
    echo [OK] .env already exists, skipping.
)

:: Load env vars from .env
for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
    set "%%A=%%B"
)

:: Open browser after short delay
start "" cmd /c "timeout /t 4 >nul & start http://localhost:5000"

:: Run the app
echo.
echo ============================================
echo  App running at http://localhost:5000
echo  Press Ctrl+C to stop.
echo ============================================
echo.
.venv\Scripts\python app.py
pause
