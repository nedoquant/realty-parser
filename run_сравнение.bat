@echo off
rem Пересобрать сводный лист «Сравнение» (формулы-ссылки на листы-источники). Без сети, секунды.
chcp 65001 >/dev/null
cd /d "%~dp0"
python cian_parser.py compare
echo.
pause
