@echo off
REM ========================================
REM نص تشغيل البرنامج على Windows
REM ========================================

cd /d "%~dp0"
python main.py

if %errorlevel% neq 0 (
    echo.
    echo [خطأ] حدث خطأ أثناء تشغيل البرنامج
    echo [Error] An error occurred while running the application
    echo.
    echo يرجى التأكد من:
    echo 1. تثبيت Python
    echo 2. تثبيت جميع المتطلبات (install_windows.bat)
    echo.
    pause
)
