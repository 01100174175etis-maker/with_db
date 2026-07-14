@echo off
REM ========================================
REM نص تثبيت البرنامج على Windows
REM ========================================
echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║   تثبيت نظام إدارة الحسابات والمبيعات                      ║
echo ║   Sales Management System - Windows Setup                  ║
echo ╚════════════════════════════════════════════════════════════╝
echo.

REM التحقق من وجود Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [خطأ] Python غير مثبت على النظام!
    echo [Error] Python is not installed on your system!
    echo.
    echo يرجى تثبيت Python من: https://www.python.org/
    pause
    exit /b 1
)

echo [✓] تم التحقق من وجود Python
echo.

REM تحديث pip
echo [*] تحديث pip...
python -m pip install --upgrade pip
echo.

REM تثبيت المتطلبات
echo [*] تثبيت المتطلبات المطلوبة...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [خطأ] فشل تثبيت المتطلبات
    pause
    exit /b 1
)

echo.
echo [✓] تم تثبيت جميع المتطلبات بنجاح!
echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║   لتشغيل البرنامج، استخدم الأمر:                          ║
echo ║   python main.py                                           ║
echo ║   أو انقر على run.bat                                      ║
echo ╚════════════════════════════════════════════════════════════╝
echo.
pause
