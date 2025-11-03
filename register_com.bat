@echo off
REM Chandra Excel COM Registration Script
REM Скрипт для регистрации COM-сервера Chandra OCR в Windows

echo ============================================================
echo Chandra Excel COM Registration
echo ============================================================
echo.

REM Проверка наличия Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python не найден в PATH
    echo Установите Python 3.10 или новее с https://www.python.org/
    pause
    exit /b 1
)

echo [OK] Python найден
python --version

echo.
echo Регистрация COM-сервера...
echo.

REM Регистрация COM-сервера
python chandra_excel_com.py --register

if errorlevel 1 (
    echo.
    echo [ERROR] Ошибка при регистрации COM-сервера
    pause
    exit /b 1
)

echo.
echo ============================================================
echo [SUCCESS] COM-сервер успешно зарегистрирован!
echo ============================================================
echo.
echo Теперь можно использовать Chandra OCR в Excel VBA:
echo   Set obj = CreateObject("ChandraExcel.Processor")
echo.
echo Для удаления регистрации запустите:
echo   python chandra_excel_com.py --unregister
echo.
pause

