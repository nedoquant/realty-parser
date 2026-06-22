# -*- coding: utf-8 -*-
"""БРАУЗЕРLESS сбор цен Авито по 16 городам — через curl_cffi (Chrome-TLS) + куки.
Никакого undetected-chromedriver: меняем город прямыми API-запросами и читаем цены.

Цепочка на город:
  coords/by_address -> js/v2/geo/position (geoFieldsHash) -> item-edit/load (version)
  -> item-edit/submit/v2 (смена города) -> getAvailableBBL/bbip-setups/vas-configurator.

Шаблон тела submit (60 полей) и заголовки берём из www.avito.ru.har (снят пользователем),
куки — из avito_cookie.txt. Пишем в лист «Avito Price» (cp.write_avito_price)."""
import sys, json, re, time, urllib.parse, datetime
sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
from curl_cffi import requests
import config
import cian_parser as cp

ITEM = config.AVITO_PRICE_ITEM
DATE = datetime.date.today().isoformat()
HAR = "www.avito.ru.har"
VF = config.AVITO_PRICE_VASFROM
BASE = config.AVITO_BASE

# Порядок ЧЕРЕДУЕТ тарифы (соседи из разных ценовых групп) — чтобы значение нового
# города заведомо отличалось от предыдущего: так схождение кэша детектируется быстро
# и однозначно (не путается с залипшим значением соседа).
ADDRS = [
    ("Москва", "Москва, Садовая-Черногрязская улица, 20/28"),
    ("Казань", "Республика Татарстан (Татарстан), Казань, улица Чернышевского, 6/2"),
    ("Новосибирск", "Новосибирск, Красный проспект, 42"),
    ("Ростов-на-Дону", "Ростов-на-Дону, проспект Чехова, 55"),
    ("Нижний Новгород", "Нижний Новгород, улица Ульянова, 12"),
    ("Воронеж", "Воронеж, Плехановская улица, 8"),
    ("Красноярск", "Красноярск, проспект Мира, 122"),
    ("Челябинск", "Челябинск, улица Елькина, 47"),
    ("Омск", "Омск, улица Ленина, 2"),
    ("Краснодар", "Краснодар, микрорайон Центральный, Длинная улица, 127"),
    ("Санкт-Петербург", "Санкт-Петербург, улица Александра Невского, 10"),
    ("Уфа", "Республика Башкортостан, Уфа, Революционная улица, 57"),
    ("Екатеринбург", "Свердловская область, Екатеринбург, улица Попова, 6"),
    ("Самара", "Самара, Молодогвардейская улица, 139"),
    ("Волгоград", "Волгоград, улица Мира, 16"),
    ("Пермь", "Пермь, Осинская улица, 13"),
]


def rub(s):
    d = re.sub(r"\D", "", str(s))
    return int(d) if d else None


def parse_template():
    """Шаблон тела submit (список (name,value)) + браузерные заголовки из HAR."""
    har = json.load(open(HAR, encoding="utf-8"))
    for e in har["log"]["entries"]:
        if "item-edit/submit/v2" in e["request"]["url"]:
            boundary = e["request"]["postData"]["mimeType"].split("boundary=")[1]
            body = e["request"]["postData"]["text"]
            fields = []
            for part in body.split("--" + boundary):
                if 'name="' not in part:
                    continue
                name = re.search(r'name="([^"]+)"', part).group(1)
                seg = part.split("\r\n\r\n", 1)
                if len(seg) < 2:
                    continue
                fields.append((name, seg[1].rsplit("\r\n", 1)[0]))
            hdrs = {hd["name"]: hd["value"] for hd in e["request"]["headers"]
                    if hd["name"].lower() in ("user-agent", "accept-language", "sec-ch-ua",
                                              "sec-ch-ua-mobile", "sec-ch-ua-platform", "dnt")}
            return fields, hdrs
    raise RuntimeError("в HAR нет item-edit/submit/v2 — пере-сними HAR")


def fees_template():
    """Шаблон тела fees/methods из HAR (для «прайминга» расчёта под город)."""
    har = json.load(open(HAR, encoding="utf-8"))
    for e in har["log"]["entries"]:
        if "fees/" in e["request"]["url"] and "methods" in e["request"]["url"]:
            return json.loads(e["request"]["postData"]["text"])
    return None


def prime_methods(s, J, geo):
    """Вызвать fees/methods под город (как fee-страница после «Сохранить») и вернуть
    ЯКОРЬ — params['201'] из ответа. Это город-корректная «Оценить эффект
    продвижения» БЕЗ лага: сервер пересчитывает её по locationId сразу. Дальше по
    этому якорю ждём, пока setups (асинхронный) догонит текущий город. None -> нет."""
    if not _FEES:
        return None
    b = json.loads(json.dumps(_FEES))
    b["locationId"] = geo["locationId"]
    b["parentLocationId"] = geo.get("parentLocationId", b.get("parentLocationId"))
    b["params"]["493"] = geo["address"]
    b["params"]["100006"] = str(geo["locationId"])
    ifp = b.get("itemFormParams", {})
    ifp.update({"address": geo["address"], "locationId": str(geo["locationId"]),
                "coords[lat]": str(geo["latitude"]), "coords[lng]": str(geo["longitude"]),
                "geoFieldsHash": geo["geoFieldsHash"],
                "district_id": str(geo.get("districtId", "") or ""),
                "metro_id": str(geo.get("metroId", "") or "")})
    try:
        d = s.post(f"{BASE}/web/4/fees/{ITEM}/methods", json=b, headers=J, timeout=30).json()
        return rub(d.get("priceRequestAdditionalParams", {}).get("params", {}).get("201"))
    except Exception:  # noqa: BLE001
        return None


_FEES = None  # шаблон тела fees/methods (для прайминга), грузится в make_session


def make_session():
    global _FEES
    _FEES = fees_template()
    tmpl, base_h = parse_template()
    raw = open(config.AVITO_PRICE_COOKIE, encoding="utf-8").read().strip()
    cookies = {k.strip(): v.strip() for k, v in
               (p.split("=", 1) for p in raw.split(";") if "=" in p)}
    H = dict(base_h)
    H.update({"accept": "application/json", "origin": "https://www.avito.ru",
              "sec-fetch-dest": "empty", "sec-fetch-mode": "cors",
              "sec-fetch-site": "same-origin",
              "referer": f"https://www.avito.ru/items/edit/{ITEM}"})
    J = dict(H); J["content-type"] = "application/json"   # для JSON-запросов
    s = requests.Session(impersonate="chrome", cookies=cookies, headers=H)
    return s, H, J, tmpl


def geo_for(s, J, addr):
    """Адрес -> {geoFieldsHash, locationId, districtId, lat, lng, address}."""
    q = urllib.parse.quote(addr)
    co = s.get(f"{BASE}/web/1/coords/by_address?address={q}", headers=J, timeout=30).json()
    pt = co["point"]
    body = {"categoryId": 24, "zoom": 16, "params": {}, "itemId": ITEM,
            "latitude": pt["latitude"], "longitude": pt["longitude"],
            "address": co["normalizedAddress"], "isRadius": False,
            "eventName": "item-add-geo-text-suggest"}
    g = s.post(f"{BASE}/js/v2/geo/position", json=body, headers=J, timeout=30).json()
    return g


def current_version(s, J):
    ld = s.post(f"{BASE}/item-edit/load/v2?", json={"itemId": ITEM, "fromItemId": 0},
                headers=J, timeout=30).text
    m = re.search(r'"version"\s*:\s*(\d+)', ld)
    return m.group(1) if m else None


def set_city(s, H, J, tmpl, geo):
    """Сменить город объявления через submit/v2. True при успехе."""
    ver = current_version(s, J)
    swap = {
        "params[493]": geo["address"], "address": geo["address"],
        "locationId": str(geo["locationId"]), "geoFieldsHash": geo["geoFieldsHash"],
        "coords[lat]": str(geo["latitude"]), "coords[lng]": str(geo["longitude"]),
        "district_id": str(geo.get("districtId", "") or ""),
        "metro_id": str(geo.get("metroId", "") or ""),
    }
    if ver:
        swap["version"] = ver
    bnd = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
    parts = [f"--{bnd}\r\nContent-Disposition: form-data; name=\"{n}\"\r\n\r\n"
             f"{swap.get(n, v)}" for n, v in tmpl]
    body = ("\r\n".join(parts) + f"\r\n--{bnd}--\r\n").encode("utf-8")
    HS = dict(H); HS["content-type"] = f"multipart/form-data; boundary={bnd}"
    r = s.post(f"{BASE}/item-edit/submit/v2", data=body, headers=HS, timeout=40)
    try:
        return r.json().get("status") == "ok"
    except Exception:  # noqa: BLE001
        return False


def get_pub(s, J):
    d = s.post(f"{BASE}/web/2/bbl-api/getAvailableBBL", json={"itemId": ITEM},
               headers=J, timeout=30).json()
    by = {c.get("ttl"): c.get("priceOrigin") for c in d.get("configs", [])}
    return (by.get(14), by.get(30), by.get(60))


def get_promo(s, J):
    d = s.post(f"{BASE}/web/1/bbip/private/setups",
               json={"itemId": ITEM, "vasFrom": VF, "s": ["bbl"]}, headers=J, timeout=30).json()
    return [(p.get("title", "?"), rub(p.get("oldPriceFormatted")))
            for p in d.get("data", {}).get("presets", [])]


def get_highlight(s, J):
    d = s.post(f"{BASE}/web/1/vas/configurator",
               json={"itemId": ITEM, "vasFrom": VF, "s": ["bbl"],
                     "referer": f"{BASE}/bbl/{ITEM}/period"}, headers=J, timeout=30).json()
    out = []
    for w in d.get("data", {}).get("widgets", []):
        nm = w.get("name") or (w.get("data") or {}).get("name")
        op = w.get("oldPrice")
        if op is None:
            op = (w.get("data") or {}).get("oldPrice")
        if nm and op is not None:
            out.append((nm, rub(op)))
    return out


def _promo_first(promo):
    """Цена пресета «Оценить эффект продвижения» (его сверяем с якорем 201)."""
    for t, p in promo:
        if t.startswith("Оцен"):
            return p
    return promo[0][1] if promo else None


def valid_triple(triple):
    """Тройка пресетов одного города связана пропорцией: 2-й ≈ 1.93-1.97×1-го,
    3-й ≈ 3.88-3.97×1-го. Проверяем, что это валидная тройка (не склейка)."""
    try:
        a, b, c = triple
        if not a:
            return False
        return 1.88 <= b / a <= 2.03 and 3.80 <= c / a <= 4.08
    except (TypeError, ZeroDivisionError):
        return False


DEFAULT_PROMO = (1365, 2645, 5350)   # ДЕФОЛТ объявления (config.oldPrice 27300=«1365 ₽»),
                                     # показывается пока город-прогноз НЕ досчитался — не город!


def read_promo_converge(s, J, base, resubmit=None, win=8, budget=720, step=5, repush=90):
    """Кэш setups забит значением ПРЕДЫДУЩЕГО города (base) и сходится к истинному
    значению текущего за минуты ТИШИНЫ. Расхождение бывает МЕДЛЕННЫМ (до ~500c!),
    поэтому НЕ срезаем по таймауту: принимаем ТОЛЬКО когда последние `win` чтений
    ОДИНАКОВЫ И отличаются от base (новый город вытеснил предыдущий). Если за budget
    так и не разошлось с base — НЕ выдаём базу как истину, а флаг (converged=False),
    пусть пометится ⚠ПРОВЕРИТЬ. (Порядок городов чередует тарифы → new!=base всегда.)
    Возвращает (promo, converged, triple)."""
    from collections import deque, Counter
    print(f"   ⏳ жду схождения != base {base} (единогласие {win}, до {budget}c)...")
    buf = deque(maxlen=win); full = {}
    t0 = time.time(); last_log = 0; last_push = 0
    while time.time() - t0 < budget:
        el = int(time.time() - t0)
        if resubmit and el - last_push >= repush:      # пере-проталкиваем город на ноды
            last_push = el
            try:
                resubmit()
            except Exception:  # noqa: BLE001
                pass
        pr = get_promo(s, J)
        if pr:
            k = tuple(p for _, p in pr); buf.append(k); full[k] = pr
            el = int(time.time() - t0)
            if el - last_log >= 60:
                last_log = el
                print(f"      {el}c  окно={Counter(buf).most_common(3)}")
            if len(buf) == win and len(set(buf)) == 1 and (base is None or buf[0] != base):
                print(f"      сошлось за {el}c -> {buf[0]}")
                return full[buf[0]], True, buf[0]
        time.sleep(step)
    c = Counter(buf)
    if not c:
        return None, False, None
    k, _ = c.most_common(1)[0]
    return full[k], False, k     # не разошлось с base за бюджет -> ⚠ПРОВЕРИТЬ


def collect(s, H, J, tmpl, city, addr):
    """Быстрый сбор: ТОЛЬКО публикация + выделение (мгновенны и точны после смены
    города). ПРОДВИЖЕНИЕ здесь НЕ берём — его кэш у Авито сходится минутами и
    нестабилен в батче, поэтому собирается ВРУЧНУЮ. Логика схождения продвижения
    сохранена в read_promo_converge/prime_methods/get_promo на будущее."""
    geo = geo_for(s, J, addr)
    if not geo.get("geoFieldsHash"):
        print("   ⚠ гео не получено"); return None
    if not set_city(s, H, J, tmpl, geo):
        print("   ⚠ submit не ok"); return None
    pub = get_pub(s, J)
    hl = get_highlight(s, J)
    # ВАЖНО: отдаём ВСЕ 8 значений в каноническом порядке (write_avito_price кладёт
    # их по столбцам ПОЗИЦИОННО). Продвижение = None (3 пустых места) — заполняется
    # вручную; при записи существующее продвижение НЕ затирается (см. merge).
    rows = [("Публикация", "14 дней", pub[0]),
            ("Публикация", "30 дней", pub[1]),
            ("Публикация", "60 дней", pub[2]),
            ("Продвижение (5 дн)", "Оценить", None),
            ("Продвижение (5 дн)", "Выбирают чаще", None),
            ("Продвижение (5 дн)", "Сильнее", None)]
    for nm, p in hl:
        rows.append(("Выделение", nm, p))
    return rows


def main():
    todo = ADDRS
    if len(sys.argv) > 1:
        want = [a.lower() for a in sys.argv[1:]]
        todo = [t for t in ADDRS if any(w in t[0].lower() for w in want)]
    s, H, J, tmpl = make_session()
    print(f"сессия готова, шаблон {len(tmpl)} полей; городов: {len(todo)}")
    done, failed = [], []
    for i, (city, addr) in enumerate(todo, 1):
        print(f"\n[{i}/{len(todo)}] {city}: {addr}")
        try:
            rows = collect(s, H, J, tmpl, city, addr)
        except Exception as e:  # noqa: BLE001
            print("   ! ошибка:", str(e)[:100]); rows = None
        if not rows:
            failed.append(city); continue
        for g, n, v in rows:
            sh = f"{v:,}".replace(",", " ") if v is not None else "—"
            print(f"      {n:32} {sh:>9}")
        cp.write_avito_price([(g, n, city, v) for g, n, v in rows], DATE)
        print(f"   -> сохранено ({city})"); done.append(city)
    print(f"\n=== ГОТОВО: {len(done)}/{len(todo)} (публикация + выделение) ===")
    if failed:
        print("   НЕ удалось:", ", ".join(failed))
    print("   Продвижение собери ВРУЧНУЮ в лист «Avito Price» (3 столбца), затем")
    print("   запусти run_синхр_динамика.bat — он перенесёт всё в «Динамика цен».")


if __name__ == "__main__":
    main()
