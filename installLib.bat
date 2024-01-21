@echo off

REM Установка библиотек из файла req.txt
pip install -r req.txt

REM Проверка статуса установки
if %ERRORLEVEL% NEQ 0 (
    echo Установка библиотек завершилась с ошибкой.
    exit /b %ERRORLEVEL%
)

echo Все библиотеки успешно установлены.