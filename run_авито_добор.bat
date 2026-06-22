@echo off
rem Добор оставшихся городов Авито (запускать на ДОМАШНЕМ интернете, БЕЗ VPN!)
chcp 65001 >/dev/null
cd /d "%~dp0"
python cian_parser.py avito "Самара,Пермь,Волгоград"
echo.
pause
