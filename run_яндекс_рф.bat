@echo off
rem Яндекс: вся Россия (85 субъектов) -> лист «Яндекс регионы» + строки «Вся Россия». ~35-40 мин
chcp 65001 >/dev/null
cd /d "%~dp0"
python cian_parser.py yandex_rf
echo.
pause
