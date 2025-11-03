@echo off
REM Chandra Excel Integration Setup Script
REM Автоматическая настройка окружения

echo ============================================================
echo Chandra Excel Integration Setup
echo ============================================================
echo.

REM Проверка Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python не найден в PATH
    echo Установите Python 3.10+ с https://www.python.org/
    pause
    exit /b 1
)

echo [OK] Python найден
python --version
echo.

REM Установка зависимостей
echo Установка зависимостей...
python -m pip install --upgrade pip
pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Ошибка при установке зависимостей
    pause
    exit /b 1
)

echo.
echo [OK] Зависимости установлены
echo.

REM Настройка переменных окружения
echo Настройка переменных окружения...
echo.

if not exist "env.example" (
    echo [WARNING] Файл env.example не найден
    echo Создайте файл .env вручную
    goto :check_registration
)

echo Скопируйте env.example в .env и заполните значения
echo   copy env.example .env
echo.
echo Затем отредактируйте .env и укажите:
echo   - OPENROUTER_API_KEY (получите на https://openrouter.ai/)
echo   - OPENROUTER_MODEL (используйте openai/gpt-4o или другой)
echo.

:check_registration
REM Регистрация COM-сервера
echo Регистрация COM-сервера...
echo.

python chandra_excel_com.py --register

if errorlevel 1 (
    echo.
    echo [ERROR] Ошибка при регистрации COM-сервера
    echo Попробуйте запустить register_com.bat от имени администратора
    pause
    exit /b 1
)

echo.
echo [OK] COM-сервер зарегистрирован
echo.

REM Финальная проверка
echo ============================================================
echo Установка завершена!
echo ============================================================
echo.
echo Следующие шаги:
echo   1. Создайте файл .env и укажите OPENROUTER_API_KEY
echo   2. Откройте Excel и импортируйте макросы
echo   3. Назначьте горячую клавишу для ShowChandraForm
echo   4. Протестируйте на любом PDF файле
echo.
echo См. README_CHANDRA_EXCEL.md для подробностей
echo.
pause

