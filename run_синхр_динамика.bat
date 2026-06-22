@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Перенести текущие листы «Avito Price» и «CIAN Price» в лист «Динамика цен»
rem (добавит/обновит столбец текущего месяца + Δ%). Запускать ПОСЛЕ того, как
rem впишешь продвижение Авито вручную. Без сети, секунды. Закрой Excel.
python -c "import price_dynamics as p; p.sync_avito(); p.sync_cian()"
echo.
pause
