"""Строит collector/cities.csv из cities_raw.csv: IATA-код города, координаты, код страны.

Источник — публичный автокомплит Travelpayouts (ключ не нужен).
Ручные исправления — в cities_overrides.csv (city, iata, lat, lon, note).

Запуск:  python -m collector.build_cities
"""
from __future__ import annotations

import csv
from pathlib import Path

from .common import PoliteSession, log, setup_logging

HERE = Path(__file__).resolve().parent
RAW = HERE / "cities_raw.csv"
OVERRIDES = HERE / "cities_overrides.csv"
OUT = HERE / "cities.csv"
PLACES = "https://autocomplete.travelpayouts.com/places2"
NOMINATIM = "https://nominatim.openstreetmap.org/search"

FIELDS = ["city", "country", "region", "abroad", "country_code", "iata", "lat", "lon", "match", "currency_hint"]


def lookup(session: PoliteSession, city: str, country: str) -> dict | None:
    resp = session.get(PLACES, params={"term": city, "locale": "ru", "types[]": "city"})
    resp.raise_for_status()
    items = [i for i in resp.json() if i.get("type") == "city"]
    if not items:
        return None
    want = country.lower()
    same_country = [i for i in items if want and want[:5] in (i.get("country_name") or "").lower()]
    return (same_country or items)[0]


def geocode_osm(session: PoliteSession, city: str, country: str, region: str) -> tuple | None:
    """Координаты города через Nominatim (OpenStreetMap). Правила OSM: не чаще 1 запроса в секунду."""
    name = city.split(",")[0].replace(" р-н", " район")
    items = []
    for parts in ((name, region, country), (name, country)):
        query = ", ".join(p for p in parts if p)
        resp = session.get(NOMINATIM, params={"q": query, "format": "json", "limit": 1, "addressdetails": 1, "accept-language": "ru"})
        resp.raise_for_status()
        items = resp.json()
        if items:
            break
    if not items:
        return None
    it = items[0]
    return round(float(it["lat"]), 5), round(float(it["lon"]), 5), it.get("address", {}).get("country_code", "").upper()


def main() -> None:
    setup_logging()
    overrides = {}
    if OVERRIDES.exists():
        with OVERRIDES.open(encoding="utf-8") as f:
            overrides = {r["city"]: r for r in csv.DictReader(f)}

    session = PoliteSession(0.3, 0.6)
    osm = PoliteSession(1.1, 1.5)
    osm.headers["User-Agent"] = "no-rest-boleto/0.1 (github.com/Legpopov/no-rest-boleto)"
    out = []
    with RAW.open(encoding="utf-8") as f:
        for row in csv.DictReader(f):
            rec = {k: row.get(k, "") for k in FIELDS}
            if row["city"] in overrides:
                ov = overrides[row["city"]]
                rec.update(iata=ov["iata"], lat=ov["lat"], lon=ov["lon"], match="override")
                if not rec["lat"]:
                    geo = geocode_osm(osm, row["city"], row["country"], row["region"])
                    if geo:
                        rec.update(lat=geo[0], lon=geo[1], country_code=geo[2])
            else:
                hit = lookup(session, row["city"], row["country"])
                if hit:
                    exact = hit["name"].lower().replace("ё", "е") == row["city"].lower().replace("ё", "е")
                    rec.update(
                        iata=hit["code"],
                        country_code=hit.get("country_code", ""),
                        lat=hit["coordinates"]["lat"],
                        lon=hit["coordinates"]["lon"],
                        match="exact" if exact else f"fuzzy:{hit['name']}",
                    )
                else:
                    rec["match"] = "not_found"
            log.info("%-28s %-4s %s", row["city"], rec["iata"], rec["match"])
            out.append(rec)

    with OUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(out)
    bad = [r for r in out if r["match"] != "exact" and r["match"] != "override"]
    log.info("Готово: %d городов, требуют проверки: %d", len(out), len(bad))


if __name__ == "__main__":
    main()
