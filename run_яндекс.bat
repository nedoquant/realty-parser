@echo off
rem Яндекс.Недвижимость — количество объявлений по 16 городам (лист «Яндекс»)
chcp 65001 >/dev/null
cd /d "%~dp0"
python cian_parser.py yandex
echo.
pause
