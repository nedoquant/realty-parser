@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem ПОЛНОЕ ОБНОВЛЕНИЕ всех источников по очереди (последовательно, безопасно).
rem Все парсеры пишут в один cian_объявления.xlsx, поэтому идут друг за другом.
rem Перед стартом обнови cookie: cookie.txt (CIAN) и avito_cookie.txt (цены Авито).
rem Время ~1-1.5 ч. Цены Авито = только публикация+выделение (продвижение вручную).
echo ВНИМАНИЕ: обнови cookie.txt и avito_cookie.txt перед запуском.
echo Полное обновление займёт ~1-1.5 ч. Закрой Excel с cian_объявления.xlsx.
echo Продвижение Авито собери вручную ПОСЛЕ и запусти run_синхр_динамика.bat.
pause

echo === [1/7] CIAN — количество ===
python cian_parser.py cities

echo === [2/7] CIAN — цены ===
python cian_parser.py prices
python -c "import price_dynamics; price_dynamics.sync_cian()"

echo === [3/7] Яндекс — количество ===
python cian_parser.py yandex

echo === [4/7] Авито — количество ===
python cian_parser.py avito

echo === [5/7] Домклик — количество ===
python cian_parser.py domclick

echo === [6/7] Авито — ЦЕНЫ (публикация+выделение, ~3-5 мин; продвижение вручную) ===
python avito_price_api.py

echo === [7/7] Сводный лист «Сравнение» ===
python cian_parser.py compare

echo.
echo === ВСЁ ОБНОВЛЕНО ===
pause
