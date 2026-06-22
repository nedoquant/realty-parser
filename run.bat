@echo off
rem Сбор по 16 городам (быстро, ~8 мин). Для всей РФ — run_Россия.bat, для обоих — run_всё.bat
chcp 65001 >nul
cd /d "%~dp0"
python cian_parser.py cities
echo.
pause
