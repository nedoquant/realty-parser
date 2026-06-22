# -*- coding: utf-8 -*-
"""
Конфигурация парсера ЦИАН.

Парсим ТОЛЬКО количество объявлений ("Найдено N объявлений") по матрице
КАТЕГОРИЯ × ГОРОД. Уровень — город (не область).

------------------------------------------------------------------------------
CITIES — города (топ по численности). Каждая запись:
    ("Название", "поддомен", код_региона)
  * поддомен — домен ЦИАН для города (напр. spb -> spb.cian.ru; Москва -> www)
  * код_региона — внутренний код ЦИАН (нужен только для коворкинга через cat.php)
  Коды получены из самих страниц ЦИАН (jsonQuery), проверены 2026-06-10.

CATEGORIES — категории. Каждая запись:
    ("Группа", "Название", вид, значение)
  вид = "slug"   -> URL = https://<поддомен>.cian.ru/<значение>/
  вид = "catphp" -> URL = https://www.cian.ru/cat.php?<значение>&region=<код>
                    (используется для коворкинга — у него нет «красивого» URL)
------------------------------------------------------------------------------
"""

OUTPUT_FILE = "cian_объявления.xlsx"
SHEET_NAME = "ЦИАН"
REGIONS_SHEET = "Вся Россия (регионы)"  # лист с разбивкой по 85 субъектам РФ

# Пауза между запросами, сек (защита от капчи). 336 запросов ~ 7-10 минут.
REQUEST_DELAY = 1.0

CITIES = [
    ("Москва",          "www",          1),
    ("Санкт-Петербург", "spb",          2),
    ("Новосибирск",     "novosibirsk",  4897),
    ("Екатеринбург",    "ekb",          4743),
    ("Казань",          "kazan",        4777),
    ("Красноярск",      "krasnoyarsk",  4827),
    ("Нижний Новгород", "nn",           4885),
    ("Челябинск",       "chelyabinsk",  5048),
    ("Уфа",             "ufa",          176245),
    ("Краснодар",       "krasnodar",    4820),
    ("Самара",          "samara",       4966),
    ("Ростов-на-Дону",  "rostov",       4959),
    ("Омск",            "omsk",         4914),
    ("Воронеж",         "voronezh",     4713),
    ("Пермь",           "perm",         4927),
    ("Волгоград",       "volgograd",    4704),
]

# office_type[0]=11 — коворкинг (нет красивого URL, только cat.php)
COWORKING_CATPHP = ("deal_type=rent&engine_version=2&offer_type=offices"
                    "&office_type%5B0%5D=11")

CATEGORIES = [
    # ---------- Аренда квартир ----------
    ("Аренда квартир", "Длительная: Квартиры",        "slug", "snyat-kvartiru"),
    ("Аренда квартир", "Длительная: Комнаты",          "slug", "snyat-komnatu"),
    ("Аренда квартир", "Длительная: Дома и коттеджи",  "slug", "snyat-dom"),
    ("Аренда квартир", "Посуточная аренда",            "slug", "snyat-kvartiru-posutochno"),

    # ---------- Продажи квартир ----------
    ("Продажи квартир", "Квартиры в новостройках",     "slug", "kupit-kvartiru-novostroyki"),
    ("Продажи квартир", "Квартиры во вторичке",        "slug", "kupit-kvartiru-vtorichka"),
    ("Продажи квартир", "Комнаты",                     "slug", "kupit-komnatu"),
    ("Продажи квартир", "Дома и коттеджи",             "slug", "kupit-dom"),
    ("Продажи квартир", "Участки",                     "slug", "kupit-zemelniy-uchastok"),

    # ---------- Дома и участки ----------
    ("Дома и участки", "Продажа таунхаусов",           "slug", "kupit-taunhaus"),

    # ---------- Коммерческая: аренда ----------
    ("Коммерческая · Аренда", "Офис",                  "slug", "snyat-ofis"),
    ("Коммерческая · Аренда", "Коворкинг",             "catphp", COWORKING_CATPHP),
    ("Коммерческая · Аренда", "Торговая площадь",      "slug", "snyat-torgovuyu-ploshad"),
    ("Коммерческая · Аренда", "Складское помещение",   "slug", "snyat-sklad"),

    # ---------- Коммерческая: продажа ----------
    ("Коммерческая · Продажа", "Офис",                 "slug", "kupit-ofis"),
    ("Коммерческая · Продажа", "Торговая площадь",     "slug", "kupit-torgovuyu-ploshad"),
    ("Коммерческая · Продажа", "Складское помещение",  "slug", "kupit-sklad"),
    ("Коммерческая · Продажа", "Бизнес (готовый)",     "slug", "kupit-gotoviy-biznes"),

    # ---------- Коммерческая: спец-объекты ----------
    ("Коммерческая · Спец-объекты", "Бизнес-центры",       "slug", "snyat-pomeshenie-v-biznes-centre"),
    ("Коммерческая · Спец-объекты", "Торговые центры",     "slug", "snyat-pomeshenie-v-torgovom-centre"),
    ("Коммерческая · Спец-объекты", "Складские комплексы", "slug", "snyat-pomeshenie-v-skladskom-komplekse"),
]


# =============================================================================
# ЦЕНЫ ПУБЛИКАЦИИ (лист «CIAN Price», my.cian.ru/price -> billing-tariffs API)
# =============================================================================
PRICE_SHEET = "CIAN Price"
PRICE_API = "https://api.cian.ru/billing-tariffs/v1/get-price-info/"
# Файл с cookie авторизованной сессии cian.ru (нужен — API цен требует вход).
# Обновляется вручную при протухании: DevTools -> запрос price (document)
# -> Copy as cURL -> взять строку из -b '...'. Достаточно DMIR_AUTH + cian_ruid.
COOKIE_FILE = "cookie.txt"

PRICE_DEALS = [("sale", "Продажа"), ("rent", "Аренда"), ("dailyRent", "Посуточная аренда")]
PRICE_OFFERS = [("flat", "Городская"), ("suburban", "Загородная"), ("commercial", "Коммерческая")]

# polygonId ценовой зоны для каждого города (из get-price-filters-info, 2026-06-10).
# Это ОТДЕЛЬНЫЕ коды (не region_code!), привязаны к городу.
PRICE_POLYGONS = {
    "Москва": 2000, "Санкт-Петербург": 1999, "Новосибирск": 153356,
    "Екатеринбург": 153357, "Казань": 107227, "Красноярск": 76408,
    "Нижний Новгород": 107224, "Челябинск": 153358, "Уфа": 107226,
    "Краснодар": 107222, "Самара": 107228, "Ростов-на-Дону": 107225,
    "Омск": 176074, "Воронеж": 76407, "Пермь": 76412, "Волгоград": 107529,
}


# =============================================================================
# ЯНДЕКС.НЕДВИЖИМОСТЬ (лист «Яндекс») — количество объявлений
# =============================================================================
# Капчи нет, число берём из <title> страницы ("— N объявлений"), кроме
# новостроек: там title = число ЖК, реальные квартиры в поле newbuildingOffersCount.
YANDEX_SHEET = "Яндекс"
YANDEX_BASE = "https://realty.yandex.ru"

# Слаги городов у Яндекса (порядок — как в CITIES). Проверено 2026-06-10.
YANDEX_CITY_SLUG = {
    "Москва": "moskva", "Санкт-Петербург": "sankt-peterburg",
    "Новосибирск": "novosibirsk", "Екатеринбург": "ekaterinburg",
    "Казань": "kazan", "Красноярск": "krasnoyarsk",
    "Нижний Новгород": "nizhniy_novgorod", "Челябинск": "chelyabinsk",
    "Уфа": "ufa", "Краснодар": "krasnodar", "Самара": "samara",
    "Ростов-на-Дону": "rostov-na-donu", "Омск": "omsk", "Воронеж": "voronezh",
    "Пермь": "perm", "Волгоград": "volgograd",
}

# (Группа, Название, слаг-пути, режим извлечения: "title" | "newbuilding")
YANDEX_CATEGORIES = [
    ("Купить", "Квартира",               "kupit/kvartira",                   "title"),
    ("Купить", "Квартира в новостройке",  "kupit/novostrojka",                "newbuilding"),
    ("Купить", "Квартира во вторичке",    "kupit/kvartira/vtorichniy-rynok",  "title"),
    ("Купить", "Комната",                 "kupit/komnata",                    "title"),
    ("Купить", "Дом или коттедж",         "kupit/dom",                        "title"),
    ("Купить", "Участок",                 "kupit/uchastok",                   "title"),
    ("Купить", "Гараж или машиноместо",   "kupit/garazh",                     "title"),

    ("Снять", "Квартира",                 "snyat/kvartira",                   "title"),
    ("Снять", "Комната",                  "snyat/komnata",                    "title"),
    ("Снять", "Дом или коттедж",          "snyat/dom",                        "title"),
    ("Снять", "Гараж или машиноместо",    "snyat/garazh",                     "title"),
    ("Снять", "Посуточно",                "snyat/kvartira/posutochno",        "title"),

    ("Коммерческая", "Продажа",           "kupit/kommercheskaya-nedvizhimost", "title"),
    ("Коммерческая", "Аренда",            "snyat/kommercheskaya-nedvizhimost", "title"),
]

# Лист с разбивкой Яндекса по всем 85 субъектам РФ (для «всей России»).
YANDEX_REGIONS_SHEET = "Яндекс регионы"

# Слаги 85 субъектов РФ у Яндекса (выверены 2026-06-11, follow-redirect).
# У Яндекса нет страны-URL, поэтому «вся Россия» = сумма по этим субъектам.
# Чувашия точным слагом не отдаётся -> прокси cheboksary (столица, ~весь объём).
YANDEX_REGION_SLUG = {
    "Адыгея": "adygeya", "Алтай": "altay", "Алтайский": "altayskiy_kray",
    "Амурская": "amurskaya_oblast", "Архангельская": "arhangelskaya_oblast",
    "Астраханская": "astrahanskaya_oblast", "Башкортостан": "bashkortostan",
    "Белгородская": "belgorodskaya_oblast", "Брянская": "bryanskaya_oblast",
    "Бурятия": "buryatiya", "Владимирская": "vladimirskaya_oblast",
    "Волгоградская": "volgogradskaya_oblast", "Вологодская": "vologodskaya_oblast",
    "Воронежская": "voronezhskaya_oblast", "Дагестан": "dagestan",
    "Еврейская": "evreyskaya_ao", "Забайкальский": "zabaykalskiy_kray",
    "Ивановская": "ivanovskaya_oblast", "Ингушетия": "ingushetiya",
    "Иркутская": "irkutskaya_oblast", "Кабардино-Балкарская": "kabardino-balkariya",
    "Калининградская": "kaliningradskaya_oblast", "Калмыкия": "kalmykiya",
    "Калужская": "kaluzhskaya_oblast", "Камчатский": "kamchatskiy_kray",
    "Карачаево-Черкесская": "karachaevo-cherkesiya", "Карелия": "kareliya",
    "Кемеровская": "kemerovskaya_oblast", "Кировская": "kirovskaya_oblast", "Коми": "komi",
    "Костромская": "kostromskaya_oblast", "Краснодарский": "krasnodarskiy_kray",
    "Красноярский": "krasnoyarskiy_kray", "Крым": "krym",
    "Курганская": "kurganskaya_oblast", "Курская": "kurskaya_oblast",
    "Ленинградская": "leningradskaya_oblast", "Липецкая": "lipetskaya_oblast",
    "Магаданская": "magadanskaya_oblast", "Марий Эл": "mariy_el", "Мордовия": "mordoviya",
    "Москва": "moskva", "Московская": "moskovskaya_oblast",
    "Мурманская": "murmanskaya_oblast", "Ненецкий": "nenetskiy_ao",
    "Нижегородская": "nizhegorodskaya", "Новгородская": "novgorodskaya_oblast",
    "Новосибирская": "novosibirskaya_oblast", "Омская": "omskaya_oblast",
    "Оренбургская": "orenburgskaya_oblast", "Орловская": "orlovskaya_oblast",
    "Пензенская": "penzenskaya_oblast", "Пермский": "permskiy_kray",
    "Приморский": "primorskiy_kray", "Псковская": "pskovskaya_oblast",
    "Ростовская": "rostovskaya_oblast", "Рязанская": "ryazanskaya_oblast",
    "Самарская": "samarskaya_oblast", "Санкт-Петербург": "sankt-peterburg",
    "Саратовская": "saratovskaya_oblast", "Саха (Якутия)": "saha_yakutiya",
    "Сахалинская": "sahalinskaya_oblast", "Свердловская": "sverdlovskaya_oblast",
    "Севастополь": "sevastopol", "Северная Осетия - Алания": "severnaya_osetiya",
    "Смоленская": "smolenskaya_oblast", "Ставропольский": "stavropolskiy_kray",
    "Тамбовская": "tambovskaya_oblast", "Татарстан": "tatarstan",
    "Тверская": "tverskaya_oblast", "Томская": "tomskaya_oblast",
    "Тульская": "tulskaya_oblast", "Тыва": "tyva", "Тюменская": "tyumenskaya_oblast",
    "Удмуртская": "udmurtiya", "Ульяновская": "ulyanovskaya_oblast",
    "Хабаровский": "habarovskiy_kray", "Хакасия": "hakasiya",
    "Ханты-Мансийский": "hanty-mansiyskiy_ao", "Челябинская": "chelyabinskaya_oblast",
    "Чеченская": "chechenskaya_respublika", "Чувашская": "cheboksary",
    "Чукотский": "chukotskiy_ao", "Ямало-Ненецкий": "yamalo-nenetskiy_ao",
    "Ярославская": "yaroslavskaya_oblast",
}


# =============================================================================
# АВИТО (лист «Авито») — количество объявлений
# =============================================================================
# Анти-бот Авито жёсткий: ПЛЕЙН-URL (/kvartiry/prodam) блокируется (firewall ~53КБ),
# а КАНОНИЧЕСКИЙ URL с суффиксом-фильтром (-ASgB...) проходит стабильно (curl --compressed).
# Суффикс кодирует КАТЕГОРИЮ (не город) — один и тот же для всех городов.
# Число: data-marker="page-title/count" / "mainCount".
AVITO_SHEET = "Авито"
AVITO_BASE = "https://www.avito.ru"
# Файл со списком прокси (по одному в строке, формат curl: http://user:pass@host:port
# или socks5://host:port). Если есть — Авито ходит через них с ротацией (обход бана по IP).
# Можно одну строку — ротирующий шлюз (новый IP на каждый запрос со стороны провайдера).
PROXY_FILE = "proxies.txt"

# --- ЦЕНЫ Авито (лист «Avito Price») ---------------------------------------
# Цены публикации/продвижения/выделения привязаны к КОНКРЕТНОМУ объявлению и его
# ГОРОДУ (агрегированных нет). Пользователь меняет город в своём объявлении, мы
# скрапим 3 private-API через headless uc с его cookie. Подробности — память avito-price.
AVITO_PRICE_SHEET = "Avito Price"
AVITO_PRICE_COOKIE = "avito_cookie.txt"   # cookie авторизации (СЕКРЕТНО)
AVITO_PRICE_ITEM = 8114115712             # id черновика-объявления пользователя
AVITO_PRICE_VASFROM = "item_edit_wait_activation_with_lf"
# Пауза-«отстаивание» (c) после смены адреса перед чтением ПРОДВИЖЕНИЯ: прогноз —
# асинхронный, при быстром переборе городов кэш-узлы Авито засоряются и значение
# скачет; пауза даёт ему осесть до верного (как при ручной правке). См. avito-price.
AVITO_PRICE_REST = 90

AVITO_CITY_SLUG = {
    "Москва": "moskva", "Санкт-Петербург": "sankt-peterburg",
    "Новосибирск": "novosibirsk", "Екатеринбург": "ekaterinburg",
    "Казань": "kazan", "Красноярск": "krasnoyarsk",
    "Нижний Новгород": "nizhniy_novgorod", "Челябинск": "chelyabinsk",
    "Уфа": "ufa", "Краснодар": "krasnodar", "Самара": "samara",
    "Ростов-на-Дону": "rostov-na-donu", "Омск": "omsk", "Воронеж": "voronezh",
    "Пермь": "perm", "Волгоград": "volgograd",
}

# (Группа=сделка, Название=тип, путь категории с каноническим суффиксом). 2026-06-11.
# Гаражи пока не включены (не удалось достать полный суффикс) — добавить позже.
AVITO_CATEGORIES = [
    ("Продажа", "Квартиры",              "kvartiry/prodam-ASgBAgICAUSSA8YQ"),
    ("Продажа", "Квартиры новостройка",  "kvartiry/prodam/novostroyka-ASgBAgICAkSSA8YQ5geOUg"),
    ("Продажа", "Квартиры вторичка",     "kvartiry/prodam/vtorichka-ASgBAgICAkSSA8YQ5geMUg"),
    ("Продажа", "Комнаты",               "komnaty/prodam-ASgBAgICAUSQA7wQ"),
    ("Продажа", "Дома, дачи, коттеджи",  "doma_dachi_kottedzhi/prodam-ASgBAgICAUSUA9AQ"),
    ("Продажа", "Земельные участки",     "zemelnye_uchastki/prodam-ASgBAgICAUSWA9oQ"),
    ("Продажа", "Коммерческая",          "kommercheskaya_nedvizhimost/prodam-ASgBAgICAUSwCNJW"),
    ("Аренда",  "Квартиры",              "kvartiry/sdam-ASgBAgICAUSSA8gQ"),
    ("Аренда",  "Квартиры долгосрочно",  "kvartiry/sdam/na_dlitelnyy_srok-ASgBAgICAkSSA8gQ8AeQUg"),
    ("Аренда",  "Квартиры посуточно",    "kvartiry/sdam/posutochno/-ASgBAgICAkSSA8gQ8AeSUg"),
    ("Аренда",  "Комнаты",               "komnaty/sdam-ASgBAgICAUSQA74Q"),
    ("Аренда",  "Дома, дачи, коттеджи",  "doma_dachi_kottedzhi/sdam-ASgBAgICAUSUA9IQ"),
    ("Аренда",  "Земельные участки",     "zemelnye_uchastki/sdam-ASgBAgICAUSWA9wQ"),
    ("Аренда",  "Коммерческая",          "kommercheskaya_nedvizhimost/sdam-ASgBAgICAUSwCNRW"),
]


# =============================================================================
# ДОМКЛИК (лист «Домклик») — количество объявлений
# =============================================================================
# Анти-бот: Qrator QAuth (JS-челлендж). curl/обычный Selenium не проходят —
# нужен undetected-chromedriver (uc). Регион = поддомен города. Счётчик — в <title>.
# URL: https://{поддомен}.domclick.ru/{путь_категории}
DOMCLICK_SHEET = "Домклик"

DOMCLICK_CITY_SUB = {
    "Москва": "moskva", "Санкт-Петербург": "spb", "Новосибирск": "novosibirsk",
    "Екатеринбург": "ekaterinburg", "Казань": "kazan", "Красноярск": "krasnoyarsk",
    "Нижний Новгород": "nn", "Челябинск": "chelyabinsk", "Уфа": "ufa",
    "Краснодар": "krasnodar", "Самара": "samara", "Ростов-на-Дону": "rostov-na-donu",
    "Омск": "omsk", "Воронеж": "voronezh", "Пермь": "perm", "Волгоград": "volgograd",
}

# «Вся Россия» по Домклику: счётчик задаётся ПОДДОМЕНОМ-городом (кука региона
# игнорируется, поддоменов уровня области нет). Листинги городов непересекающиеся,
# поэтому сумма по ВСЕМ городам-поддоменам Домклика = вся Россия. Список вытащен
# из футера/региональных ссылок domclick.ru (2026-06-11), служебные поддомены убраны.
DOMCLICK_RUSSIA_SUBS = [
    "abakan", "anadyr", "arxangelsk", "astraxan", "balashixa", "barnaul", "belgorod",
    "birobidzhan", "blagoveshhensk", "bratsk", "bryansk", "cheboksary", "chelyabinsk",
    "cherepovec", "cherkessk", "chita", "dolgoprudnyj", "domodedovo", "dzerzhinsk",
    "ekaterinburg", "elektrostal", "elista", "engels", "gorno-altajsk", "groznyj",
    "irkutsk", "ivanovo", "izhevsk", "joshkar-ola", "kaliningrad", "kaluga", "kazan",
    "kemerovo", "kirov", "kolomna", "komsomolsk-na-amure", "korolyov", "kostroma",
    "krasnodar", "krasnogorsk", "krasnoyarsk", "kurgan", "kursk", "kyzyl", "lipeck",
    "lyubercy", "magadan", "magnitogorsk", "majkop", "maxachkala", "moskva", "murmansk",
    "mytishhi", "naberezhnye-chelny", "nalchik", "naryan-mar", "nazran", "nefteyugansk",
    "nizhnevartovsk", "nizhnij-tagil", "nn", "noginsk", "norilsk", "novokuzneck",
    "novorossijsk", "novosibirsk", "novyj-urengoj", "obninsk", "odincovo", "omsk",
    "orenburg", "orexovo-zuevo", "oryol", "penza", "perm", "petropavlovsk-kamchatskij",
    "petrozavodsk", "podolsk", "prokopevsk", "pskov", "pushkino", "ramenskoe", "reutov",
    "rostov-na-donu", "ryazan", "salexard", "samara", "saransk", "saratov",
    "sergiev-posad", "serpuxov", "sevastopol", "severodvinsk", "simferopol", "smolensk",
    "sochi", "spb", "staryj-oskol", "stavropol", "sterlitamak", "surgut", "syktyvkar",
    "tambov", "tolyatti", "tomsk", "tula", "tver", "tyumen", "ufa", "ulan-ude",
    "ulyanovsk", "velikij-novgorod", "vladikavkaz", "vladimir", "vladivostok",
    "volgograd", "vologda", "volzhskij", "voronezh", "xabarovsk", "xanty-mansijsk",
    "ximki", "yakutsk", "yaroslavl", "yuzhno-saxalinsk", "zhukovskij",
]

# Лист с расшифровкой городов, по которым считается «вся Россия» Домклика.
DOMCLICK_CITIES_SHEET = "Домклик города (РФ)"
DOMCLICK_SUB_NAMES = {
    "abakan": "Абакан", "anadyr": "Анадырь", "arxangelsk": "Архангельск",
    "astraxan": "Астрахань", "balashixa": "Балашиха", "barnaul": "Барнаул",
    "belgorod": "Белгород", "birobidzhan": "Биробиджан", "blagoveshhensk": "Благовещенск",
    "bratsk": "Братск", "bryansk": "Брянск", "cheboksary": "Чебоксары",
    "chelyabinsk": "Челябинск", "cherepovec": "Череповец", "cherkessk": "Черкесск",
    "chita": "Чита", "dolgoprudnyj": "Долгопрудный", "domodedovo": "Домодедово",
    "dzerzhinsk": "Дзержинск", "ekaterinburg": "Екатеринбург", "elektrostal": "Электросталь",
    "elista": "Элиста", "engels": "Энгельс", "gorno-altajsk": "Горно-Алтайск",
    "groznyj": "Грозный", "irkutsk": "Иркутск", "ivanovo": "Иваново", "izhevsk": "Ижевск",
    "joshkar-ola": "Йошкар-Ола", "kaliningrad": "Калининград", "kaluga": "Калуга",
    "kazan": "Казань", "kemerovo": "Кемерово", "kirov": "Киров", "kolomna": "Коломна",
    "komsomolsk-na-amure": "Комсомольск-на-Амуре", "korolyov": "Королёв",
    "kostroma": "Кострома", "krasnodar": "Краснодар", "krasnogorsk": "Красногорск",
    "krasnoyarsk": "Красноярск", "kurgan": "Курган", "kursk": "Курск", "kyzyl": "Кызыл",
    "lipeck": "Липецк", "lyubercy": "Люберцы", "magadan": "Магадан",
    "magnitogorsk": "Магнитогорск", "majkop": "Майкоп", "maxachkala": "Махачкала",
    "moskva": "Москва", "murmansk": "Мурманск", "mytishhi": "Мытищи",
    "naberezhnye-chelny": "Набережные Челны", "nalchik": "Нальчик",
    "naryan-mar": "Нарьян-Мар", "nazran": "Назрань", "nefteyugansk": "Нефтеюганск",
    "nizhnevartovsk": "Нижневартовск", "nizhnij-tagil": "Нижний Тагил",
    "nn": "Нижний Новгород", "noginsk": "Ногинск", "norilsk": "Норильск",
    "novokuzneck": "Новокузнецк", "novorossijsk": "Новороссийск",
    "novosibirsk": "Новосибирск", "novyj-urengoj": "Новый Уренгой", "obninsk": "Обнинск",
    "odincovo": "Одинцово", "omsk": "Омск", "orenburg": "Оренбург",
    "orexovo-zuevo": "Орехово-Зуево", "oryol": "Орёл", "penza": "Пенза", "perm": "Пермь",
    "petropavlovsk-kamchatskij": "Петропавловск-Камчатский", "petrozavodsk": "Петрозаводск",
    "podolsk": "Подольск", "prokopevsk": "Прокопьевск", "pskov": "Псков",
    "pushkino": "Пушкино", "ramenskoe": "Раменское", "reutov": "Реутов",
    "rostov-na-donu": "Ростов-на-Дону", "ryazan": "Рязань", "salexard": "Салехард",
    "samara": "Самара", "saransk": "Саранск", "saratov": "Саратов",
    "sergiev-posad": "Сергиев Посад", "serpuxov": "Серпухов", "sevastopol": "Севастополь",
    "severodvinsk": "Северодвинск", "simferopol": "Симферополь", "smolensk": "Смоленск",
    "sochi": "Сочи", "spb": "Санкт-Петербург", "staryj-oskol": "Старый Оскол",
    "stavropol": "Ставрополь", "sterlitamak": "Стерлитамак", "surgut": "Сургут",
    "syktyvkar": "Сыктывкар", "tambov": "Тамбов", "tolyatti": "Тольятти", "tomsk": "Томск",
    "tula": "Тула", "tver": "Тверь", "tyumen": "Тюмень", "ufa": "Уфа", "ulan-ude": "Улан-Удэ",
    "ulyanovsk": "Ульяновск", "velikij-novgorod": "Великий Новгород",
    "vladikavkaz": "Владикавказ", "vladimir": "Владимир", "vladivostok": "Владивосток",
    "volgograd": "Волгоград", "vologda": "Вологда", "volzhskij": "Волжский",
    "voronezh": "Воронеж", "xabarovsk": "Хабаровск", "xanty-mansijsk": "Ханты-Мансийск",
    "ximki": "Химки", "yakutsk": "Якутск", "yaroslavl": "Ярославль",
    "yuzhno-saxalinsk": "Южно-Сахалинск", "zhukovskij": "Жуковский",
}

# (Группа=сделка, Название, путь категории). Посуточная — отдельный сервис
# 24.domclick.ru, пока не включена.
DOMCLICK_CATEGORIES = [
    ("Покупка", "Квартиры",            "pokupka/kvartiry"),
    ("Покупка", "Квартиры вторичка",   "pokupka/kvartiry/vtorichka"),
    ("Покупка", "Квартиры новостройки","pokupka/kvartiry/novostrojki"),
    ("Покупка", "Комнаты",             "pokupka/komnaty"),
    ("Покупка", "Дом",                 "pokupka/doma"),
    ("Покупка", "Дача",                "pokupka/dachi"),
    ("Покупка", "Коттеджи",            "pokupka/kottedzh"),
    ("Покупка", "Таунхаусы",           "pokupka/taunhausa"),
    ("Покупка", "Части дома",          "pokupka/chasti-doma"),
    ("Покупка", "Участок",             "pokupka/uchastka"),
    ("Покупка", "Апартаменты",         "pokupka/apartamenty"),
    ("Покупка", "Загородная",          "pokupka/zagorodnoj-nedvizhimosti"),
    ("Покупка", "Элитная",             "pokupka/business"),
    ("Покупка", "Дуплексы",            "pokupka/dupleks"),
    ("Покупка", "Коммерческая",        "pokupka-commerce"),

    ("Аренда", "Квартиры",             "arenda/kvartiry"),
    ("Аренда", "Комнаты",              "arenda/komnaty"),
    ("Аренда", "Дома",                 "arenda/doma"),
    ("Аренда", "Дачи",                 "arenda/dachi"),
    ("Аренда", "Коттеджи",             "arenda/kottedzh"),
    ("Аренда", "Таунхаусы",            "arenda/taunhausa"),
    ("Аренда", "Части дома",           "arenda/chasti-doma"),
    ("Аренда", "Участки",              "arenda/uchastka"),
    ("Аренда", "Апартаменты",          "arenda/apartamenty"),
    ("Аренда", "Загородная",           "arenda/zagorodnoj-nedvizhimosti"),
    ("Аренда", "Коммерческая",         "arenda-commerce"),
]


# =============================================================================
# СВОДНЫЙ ЛИСТ СРАВНЕНИЯ (лист «Сравнение») — ЦИАН / Яндекс / Авито / Домклик
# =============================================================================
# Источники сравниваются по 16 городам. У каждого своя таксономия, поэтому ниже
# карта: каноническая категория -> ключ (Группа, Название) в листе источника,
# либо СПИСОК ключей (суммируем), либо None (нет такой категории у источника).
COMPARE_SHEET = "Сравнение"
COMPARE_SOURCES = ["ЦИАН", "Яндекс", "Авито", "Домклик"]

# Сокращения для читаемости карты
_C = "ЦИАН"; _Y = "Яндекс"; _A = "Авито"; _D = "Домклик"

COMPARE_ROWS = [
    ("Продажа", "Квартиры — всего", {
        _C: [("Продажи квартир", "Квартиры в новостройках"),
             ("Продажи квартир", "Квартиры во вторичке")],
        _Y: [("Купить", "Квартира в новостройке"), ("Купить", "Квартира во вторичке")],
        _A: ("Продажа", "Квартиры"),
        _D: ("Покупка", "Квартиры")}),
    ("Продажа", "Квартиры — новостройки", {
        _C: ("Продажи квартир", "Квартиры в новостройках"),
        _Y: ("Купить", "Квартира в новостройке"),
        _A: ("Продажа", "Квартиры новостройка"),
        _D: ("Покупка", "Квартиры новостройки")}),
    ("Продажа", "Квартиры — вторичка", {
        _C: ("Продажи квартир", "Квартиры во вторичке"),
        _Y: ("Купить", "Квартира во вторичке"),
        _A: ("Продажа", "Квартиры вторичка"),
        _D: ("Покупка", "Квартиры вторичка")}),
    ("Продажа", "Комнаты", {
        _C: ("Продажи квартир", "Комнаты"), _Y: ("Купить", "Комната"),
        _A: ("Продажа", "Комнаты"), _D: ("Покупка", "Комнаты")}),
    ("Продажа", "Дома и коттеджи", {
        _C: ("Продажи квартир", "Дома и коттеджи"), _Y: ("Купить", "Дом или коттедж"),
        _A: ("Продажа", "Дома, дачи, коттеджи"), _D: ("Покупка", "Дом")}),
    ("Продажа", "Участки", {
        _C: ("Продажи квартир", "Участки"), _Y: ("Купить", "Участок"),
        _A: ("Продажа", "Земельные участки"), _D: ("Покупка", "Участок")}),
    ("Продажа", "Коммерческая", {
        _C: [("Коммерческая · Продажа", "Офис"), ("Коммерческая · Продажа", "Торговая площадь"),
             ("Коммерческая · Продажа", "Складское помещение"),
             ("Коммерческая · Продажа", "Бизнес (готовый)")],
        _Y: ("Коммерческая", "Продажа"),
        _A: ("Продажа", "Коммерческая"), _D: ("Покупка", "Коммерческая")}),

    ("Аренда", "Квартиры", {
        _C: ("Аренда квартир", "Длительная: Квартиры"), _Y: ("Снять", "Квартира"),
        _A: ("Аренда", "Квартиры долгосрочно"), _D: ("Аренда", "Квартиры")}),
    ("Аренда", "Комнаты", {
        _C: ("Аренда квартир", "Длительная: Комнаты"), _Y: ("Снять", "Комната"),
        _A: ("Аренда", "Комнаты"), _D: ("Аренда", "Комнаты")}),
    ("Аренда", "Дома и коттеджи", {
        _C: ("Аренда квартир", "Длительная: Дома и коттеджи"), _Y: ("Снять", "Дом или коттедж"),
        _A: ("Аренда", "Дома, дачи, коттеджи"), _D: ("Аренда", "Дома")}),
    ("Аренда", "Посуточная", {
        _C: ("Аренда квартир", "Посуточная аренда"), _Y: ("Снять", "Посуточно"),
        _A: ("Аренда", "Квартиры посуточно"), _D: None}),
    ("Аренда", "Коммерческая", {
        _C: [("Коммерческая · Аренда", "Офис"), ("Коммерческая · Аренда", "Коворкинг"),
             ("Коммерческая · Аренда", "Торговая площадь"),
             ("Коммерческая · Аренда", "Складское помещение")],
        _Y: ("Коммерческая", "Аренда"),
        _A: ("Аренда", "Коммерческая"), _D: ("Аренда", "Коммерческая")}),
]


def build_url(kind, value, subdomain, region_code):
    """Строит URL страницы ЦИАН для категории и города."""
    if kind == "slug":
        return f"https://{subdomain}.cian.ru/{value}/"
    if kind == "catphp":
        return f"https://www.cian.ru/cat.php?{value}&region={region_code}"
    raise ValueError(f"неизвестный вид категории: {kind}")


# =============================================================================
# ДАННЫЕ ПО ВСЕЙ РОССИИ (через JSON-API api.cian.ru)
# =============================================================================
# Подпись псевдо-«города» для строки итога по стране.
RUSSIA_LABEL = "🇷🇺 Вся Россия"

# Эндпоинт со списком 85 субъектов РФ (нужен для типов аренды — см. ниже).
REGIONS_URL = "https://www.cian.ru/cian-api/site/v1/get-regions/"

# Эндпоинт поиска: POST {"jsonQuery": {...}} -> data.offerCount
SEARCH_API = "https://api.cian.ru/search-offers/v2/search-offers-desktop/"

# Для этих _type запрос «без региона» возвращает 0/мусор (проверено),
# поэтому всю РФ по ним считаем СУММОЙ по 85 субъектам. Для остальных
# типов хватает одного запроса «без региона» (сверено: расхождение 0.0%).
RENT_SUM_TYPES = {"flatrent", "suburbanrent"}


def _term(v):
    return {"type": "term", "value": v}


def _terms(v):
    return {"type": "terms", "value": v}


_EV = _term(2)  # engine_version=2

# Фильтры (jsonQuery) для API по каждой категории — без поля region.
# Ключ — (Группа, Название) как в CATEGORIES. Спец-объекты не включены
# (у них нет чистого jsonQuery), для них «Вся Россия» пока не считается.
API_QUERY = {
    ("Аренда квартир", "Длительная: Квартиры"):
        {"_type": "flatrent", "engine_version": _EV, "for_day": _term("!1")},
    ("Аренда квартир", "Длительная: Комнаты"):
        {"_type": "flatrent", "engine_version": _EV, "for_day": _term("!1"), "room": _terms([0])},
    ("Аренда квартир", "Длительная: Дома и коттеджи"):
        {"_type": "suburbanrent", "engine_version": _EV, "for_day": _term("!1"), "object_type": _terms([1])},
    ("Аренда квартир", "Посуточная аренда"):
        {"_type": "flatrent", "engine_version": _EV, "for_day": _term("1")},

    ("Продажи квартир", "Квартиры в новостройках"):
        {"_type": "flatsale", "engine_version": _EV, "building_status": _term(2)},
    ("Продажи квартир", "Квартиры во вторичке"):
        {"_type": "flatsale", "engine_version": _EV, "building_status": _term(1)},
    ("Продажи квартир", "Комнаты"):
        {"_type": "flatsale", "engine_version": _EV, "room": _terms([0])},
    ("Продажи квартир", "Дома и коттеджи"):
        {"_type": "suburbansale", "engine_version": _EV, "object_type": _terms([1])},
    ("Продажи квартир", "Участки"):
        {"_type": "suburbansale", "engine_version": _EV, "object_type": _terms([3])},

    ("Дома и участки", "Продажа таунхаусов"):
        {"_type": "suburbansale", "engine_version": _EV, "object_type": _terms([4])},

    ("Коммерческая · Аренда", "Офис"):
        {"_type": "commercialrent", "engine_version": _EV, "office_type": _terms([1])},
    ("Коммерческая · Аренда", "Коворкинг"):
        {"_type": "commercialrent", "engine_version": _EV, "office_type": _terms([11])},
    ("Коммерческая · Аренда", "Торговая площадь"):
        {"_type": "commercialrent", "engine_version": _EV, "office_type": _terms([2])},
    ("Коммерческая · Аренда", "Складское помещение"):
        {"_type": "commercialrent", "engine_version": _EV, "office_type": _terms([3])},

    ("Коммерческая · Продажа", "Офис"):
        {"_type": "commercialsale", "engine_version": _EV, "office_type": _terms([1])},
    ("Коммерческая · Продажа", "Торговая площадь"):
        {"_type": "commercialsale", "engine_version": _EV, "office_type": _terms([2])},
    ("Коммерческая · Продажа", "Складское помещение"):
        {"_type": "commercialsale", "engine_version": _EV, "office_type": _terms([3])},
    ("Коммерческая · Продажа", "Бизнес (готовый)"):
        {"_type": "commercialsale", "engine_version": _EV,
         "office_type": _terms([10]), "ready_business_types": _terms([1, 2])},
}
