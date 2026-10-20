"""Основной сбор: авиа, поезда, отели, курсы -> data/prices.json (+ data/hotels/<id>.json для карты).

prices.json содержит только рыночные данные (статистику и сырые списки цен) —
никаких правил отбора. Отбор и расчёт итоговых значений делает потребитель (Excel, сайт).

Запуск:
  python -m collector.collect                       # всё
  python -m collector.collect --sources air,fx --limit 5
  python -m collector.collect --cities "Казань,Сидней"
"""
from __future__ import annotations

import argparse
import csv
import json
import statistics as st
from datetime import datetime, timezone
from pathlib import Path

from .common import DATA_DIR, PoliteSession, env, log, setup_logging
from .sources import aviasales, cbr, ostrovok, tutu

HERE = Path(__file__).resolve().parent
ALL_SOURCES = ("fx", "air", "rail", "hotels")


def quantiles(values: list[float]) -> dict | None:
    if not values:
        return None
    v = sorted(values)

    def q(p: float) -> float:
        # линейная интерполяция, как PERCENTILE.INC в Excel
        k = (len(v) - 1) * p
        lo, hi = int(k), min(int(k) + 1, len(v) - 1)
        return round(v[lo] + (v[hi] - v[lo]) * (k - lo))

    return {"n": len(v), "min": v[0], "p25": q(0.25), "median": q(0.5), "mean": round(st.mean(v)), "p75": q(0.75), "p90": q(0.9), "max": v[-1]}


def load_cities() -> list[dict]:
    with (HERE / "cities.csv").open(encoding="utf-8") as f:
        cities = list(csv.DictReader(f))
    codes = {}
    if (HERE / "codes.csv").exists():
        with (HERE / "codes.csv").open(encoding="utf-8") as f:
            codes = {(r["city"], r["country"]): r for r in csv.DictReader(f)}
    for i, c in enumerate(cities):
        c["id"] = f"c{i:03d}"
        c.update({k: v for k, v in (codes.get((c["city"], c["country"])) or {}).items() if k not in ("city", "country")})
    return cities


# ---------- источники ----------

def collect_air(s: PoliteSession, c: dict, months: list[str]) -> dict:
    if not c.get("iata"):
        return {"status": "no_airport"}
    offers = aviasales.fetch_round_trips(s, "MOW", c["iata"], months)
    # поездки с хотя бы одной ночёвкой и не длиннее 30 дней
    trips = [o for o in offers if o["return"] and o["depart"] and 1 <= (datetime.fromisoformat(o["return"]) - datetime.fromisoformat(o["depart"])).days <= 30]
    if not trips:
        return {"status": "empty", "raw_offers": len(offers)}
    direct = [o["price"] for o in trips if (o["transfers"] or 0) == 0 and (o["return_transfers"] or 0) == 0]
    le1 = [o["price"] for o in trips if (o["transfers"] or 0) <= 1 and (o["return_transfers"] or 0) <= 1]
    return {
        "status": "ok",
        "airport": c["iata"],
        "all": quantiles([o["price"] for o in trips]),
        "max1transfer": quantiles(le1),
        "direct": quantiles(direct),
        "airlines": sorted({o["airline"] for o in trips if o["airline"]}),
    }


def collect_rail(s: PoliteSession, c: dict) -> dict:
    code = c.get("express")
    if not code:
        return {"status": "no_station"}
    there = tutu.trips(s, tutu.MOSCOW, code)
    back = tutu.trips(s, code, tutu.MOSCOW)
    if not there and not back:
        return {"status": "empty"}
    # компактно: [номер, фирменный 0/1, часы, {класс: цена}]
    pack = lambda ts: [[t["train"], int(t["firm"]), t["hours"], t["classes"]] for t in ts]  # noqa: E731
    return {"status": "ok", "station": code, "there": pack(there), "back": pack(back)}


def collect_hotels(s: PoliteSession, c: dict, pages: int) -> tuple[dict, list[dict]]:
    slug = c.get("ostrovok_slug")
    if not slug:
        return {"status": "no_slug"}, []
    try:
        res = ostrovok.fetch_city(s, slug, pages)
    except ostrovok.Blocked as e:
        log.warning("ostrovok заблокировал: %s", e)
        return {"status": "blocked"}, []
    hotels = res["hotels"]
    if not hotels:
        return {"status": "empty"}, []
    kinds = sorted({h["kind"] for h in hotels if h["kind"]})
    # компактно для Excel: [цена/ночь RUB, звёзды, рейтинг, расстояние м, тип, апартаменты 0/1]
    items = [[h["price"], h["stars"], h["rating"], h["dist_m"], h["kind"], int(h["apart"])] for h in hotels]
    summary = {
        "status": "ok",
        "source": "ostrovok",
        "slug": slug,
        "total_in_city": res["total"],
        "price_all": quantiles([h["price"] for h in hotels]),
        "kinds": kinds,
        "items": items,
    }
    return summary, hotels


# ---------- основной цикл ----------

def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--sources", default=",".join(ALL_SOURCES))
    ap.add_argument("--cities", default="", help="через запятую; пусто = все")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--months", type=int, default=int(env("MONTHS_AHEAD") or 3))
    ap.add_argument("--hotel-pages", type=int, default=5)
    ap.add_argument("--out", default=str(DATA_DIR / "prices.json"))
    args = ap.parse_args()
    setup_logging()

    sources = set(args.sources.split(","))
    cities = load_cities()
    if args.cities:
        wanted = {x.strip() for x in args.cities.split(",")}
        cities = [c for c in cities if c["city"] in wanted]
    if args.limit:
        cities = cities[: args.limit]

    out_path = Path(args.out)
    # Предыдущий снимок: если источник сегодня упал, оставляем прошлые данные с их датой
    previous = {}
    if out_path.exists():
        try:
            previous = {c["id"]: c for c in json.loads(out_path.read_text(encoding="utf-8")).get("cities", [])}
        except (ValueError, KeyError):
            previous = {}

    months = aviasales.months_ahead(args.months)
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    s_air, s_rail, s_hot = PoliteSession(0.3, 0.6), PoliteSession(0.8, 1.5), PoliteSession(1.5, 3.0)
    (DATA_DIR / "hotels").mkdir(parents=True, exist_ok=True)

    result = {"generated_at": now, "months": months, "origin": {"city": "Москва", "iata": "MOW", "express": tutu.MOSCOW}}
    if "fx" in sources:
        result["fx"] = cbr.fetch_rates()

    rows = []
    for n, c in enumerate(cities, 1):
        prev = previous.get(c["id"], {})
        row = {
            "id": c["id"], "city": c["city"], "country": c["country"], "region": c["region"],
            "abroad": c["abroad"] == "1", "country_code": c["country_code"],
            "lat": float(c["lat"]) if c["lat"] else None, "lon": float(c["lon"]) if c["lon"] else None,
        }
        for src, fn in (
            ("air", lambda: collect_air(s_air, c, months)),
            ("rail", lambda: collect_rail(s_rail, c)),
            ("hotels", lambda: collect_hotels(s_hot, c, args.hotel_pages)),
        ):
            if src not in sources:
                if src in prev:
                    row[src] = prev[src]
                continue
            try:
                data = fn()
                if src == "hotels":
                    data, full = data
                    if full:
                        (DATA_DIR / "hotels" / f"{c['id']}.json").write_text(
                            json.dumps({"city": c["city"], "updated_at": now, "slug": data["slug"], "hotels": full}, ensure_ascii=False),
                            encoding="utf-8",
                        )
                data["updated_at"] = now
            except Exception as e:  # noqa: BLE001
                log.warning("%s %s: %s", src, c["city"], e)
                data = {"status": "error", "error": f"{type(e).__name__}: {str(e)[:200]}", "updated_at": now}
            # если сегодня неудача, а раньше были данные — сохраняем старые с пометкой
            if data["status"] != "ok" and prev.get(src, {}).get("status") == "ok":
                kept = dict(prev[src])
                kept["stale"] = True
                kept["last_attempt"] = {"at": now, "status": data["status"]}
                data = kept
            row[src] = data
        log.info("[%d/%d] %s: air=%s rail=%s hotels=%s", n, len(cities), c["city"],
                 row.get("air", {}).get("status"), row.get("rail", {}).get("status"), row.get("hotels", {}).get("status"))
        rows.append(row)

    # города, которые не обрабатывались в этом запуске, сохраняем из прошлого снимка
    done = {r["id"] for r in rows}
    rows += [p for pid, p in previous.items() if pid not in done]
    rows.sort(key=lambda r: r["id"])
    result["cities"] = rows
    result["coverage"] = {
        src: sum(1 for r in rows if r.get(src, {}).get("status") == "ok" and not r[src].get("stale")) for src in ("air", "rail", "hotels")
    }
    result["coverage"]["total"] = len(rows)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(result, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    log.info("Записано %s: покрытие %s", out_path, result["coverage"])


if __name__ == "__main__":
    main()
