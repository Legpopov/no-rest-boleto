# no-rest-boleto

Сборщик открытых рыночных цен на поездки: авиабилеты из Москвы туда-обратно, поезда и отели
по списку городов России и мира. Результат — `data/prices.json`, обновляется по расписанию,
и сайт с поиском вариантов поездки по датам и картой отелей.

## Источники

| Данные | Источник |
|---|---|
| Авиабилеты | Aviasales Data API (Travelpayouts) |
| Поезда | API Яндекс Расписаний |
| Отели | публичная выдача сайтов бронирования |
| Курсы валют | ЦБ РФ |
| Координаты | Travelpayouts autocomplete, OpenStreetMap Nominatim |

## Структура

```
collector/          сборщик (Python)
  cities_raw.csv    список городов
  cities_overrides.csv  ручные правки аэропортов и координат
  build_cities.py   привязка городов к IATA и координатам -> cities.csv
  sources/          модули источников
data/               результат сбора (prices.json)
site/               сайт (Netlify)
excel/              VBA-модуль импорта prices.json в Excel
```

## Запуск локально

```
pip install -r requirements.txt
python -m collector.build_cities
```

Ключи API передаются только через переменные окружения
(`TRAVELPAYOUTS_TOKEN`, `YANDEX_RASP_KEY`) и в репозиторий не попадают.
