"""Подбирает для каждого города адрес на Островке и код станции «Экспресс» -> collector/codes.csv.

Сайты отвечают только с «чистых» IP, поэтому запускается в GitHub Actions (workflow resolve-codes).
Ручные исправления: колонки ostrovok_slug / express в cities_overrides.csv имеют приоритет.

Запуск:  python -m collector.resolve_codes
"""
from __future__ import annotations

import csv
from pathlib import Path

from .common import PoliteSession, log, setup_logging
from .sources import ostrovok, tutu

HERE = Path(__file__).resolve().parent
CITIES = HERE / "cities.csv"
OUT = HERE / "codes.csv"
OVERRIDES = HERE / "codes_overrides.csv"
RAIL_COUNTRIES = {"RU", "BY", "KZ"}
# названия станций в системе «Экспресс», отличающиеся от названия города
RAIL_ALIASES = {"Орел": "Орёл"}
FIELDS = ["city", "country", "ostrovok_slug", "ostrovok_match", "currency", "express"]


def main() -> None:
    setup_logging()
    s = PoliteSession(1.0, 2.0)
    prev = {}
    if OUT.exists():
        with OUT.open(encoding="utf-8") as f:
            prev = {(r["city"], r["country"]): r for r in csv.DictReader(f)}

    overrides = {}
    if OVERRIDES.exists():
        with OVERRIDES.open(encoding="utf-8") as f:
            overrides = {r["city"]: r for r in csv.DictReader(f)}

    out = []
    with CITIES.open(encoding="utf-8") as f:
        for c in csv.DictReader(f):
            key = (c["city"], c["country"])
            rec = dict(prev.get(key) or {"city": c["city"], "country": c["country"]})
            name = c["city"].split(",")[0]
            ov = overrides.get(c["city"], {})
            if ov.get("ostrovok_slug"):
                rec.update(ostrovok_slug=ov["ostrovok_slug"], ostrovok_match="override")
            elif rec.get("ostrovok_match") != "exact":
                try:
                    hit = ostrovok.resolve_slug(s, name, c["country_code"])
                    if hit:
                        rec["ostrovok_slug"] = hit["slug"]
                        rec["ostrovok_match"] = "exact" if hit["exact"] else f"fuzzy:{hit['region_name']}"
                        rec["currency"] = hit.get("currency") or ""
                    else:
                        rec.update(ostrovok_slug="", ostrovok_match="not_found")
                except Exception as e:  # noqa: BLE001
                    rec["ostrovok_match"] = f"error:{type(e).__name__}"
            if ov.get("express"):
                rec["express"] = ov["express"]
            elif not rec.get("express") and (c["abroad"] == "0" or c["country_code"] in RAIL_COUNTRIES):
                try:
                    rec["express"] = tutu.resolve_station(s, name) or ""
                    if not rec["express"] and name in RAIL_ALIASES:
                        rec["express"] = tutu.resolve_station(s, RAIL_ALIASES[name]) or ""
                except Exception as e:  # noqa: BLE001
                    log.warning("tutu %s: %s", name, e)
            log.info("%-26s %-28s %s %s", name, rec.get("ostrovok_slug", ""), rec.get("ostrovok_match", ""), rec.get("express", ""))
            out.append({k: rec.get(k, "") for k in FIELDS})

    with OUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(out)
    log.info(
        "Готово: Островок %d/%d (точно %d), Экспресс %d",
        sum(1 for r in out if r["ostrovok_slug"]),
        len(out),
        sum(1 for r in out if r["ostrovok_match"] == "exact"),
        sum(1 for r in out if r["express"]),
    )


if __name__ == "__main__":
    main()
