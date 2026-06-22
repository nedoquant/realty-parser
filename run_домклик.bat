@echo off
rem Домклик — количество объявлений по 16 городам (лист «Домклик»). Откроется окно Chrome ~25-30 мин, НЕ закрывать!
chcp 65001 >/dev/null
cd /d "%~dp0"
python cian_parser.py domclick
echo.
pause
