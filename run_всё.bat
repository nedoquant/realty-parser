@echo off
chcp 65001 >nul
cd /d "%~dp0"
python cian_parser.py all
echo.
pause
