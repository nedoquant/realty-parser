@echo off
rem Сбор цен публикации (лист «CIAN Price»). Требует свежий cookie.txt (вход на cian).
chcp 65001 >nul
cd /d "%~dp0"
python cian_parser.py prices
python -c "import price_dynamics; price_dynamics.sync_cian()"
echo.
pause
