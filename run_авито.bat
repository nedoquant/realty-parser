@echo off
rem Авито — количество объявлений по 16 городам (лист «Авито»)
chcp 65001 >/dev/null
cd /d "%~dp0"
python cian_parser.py avito
echo.
pause
