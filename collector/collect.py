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
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from .common import DATA_DIR, PoliteSession, env, log, setup_logging
from .sources import atlas, aviasales, cbr, ostrovok, tutu

HERE = Path(__file__).resolve().parent
ALL_SOURCES = ("fx", "air", "rail", "bus", "hotels")


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
    if c.get("iata") == "MOW":
        return {"status": "origin"}
    if not c.get("iata"):
        return {"status": "no_airport"}
    try:
        offers = aviasales.fetch_round_trips(s, "MOW", c["iata"], months)
    except aviasales.NotFlightable as e:
        return {"status": "not_flightable", "airport": c["iata"], "note": str(e)[:120]}
    # поездки с хотя бы одной ночёвкой и не длиннее 30 дней
    trips = [o for o in offers if o["return"] and o["depart"] and 1 <= (datetime.fromisoformat(o["return"]) - datetime.fromisoformat(o["depart"])).days <= 30]
    if not trips:
        return {"status": "empty", "raw_offers": len(offers)}
    # по неделям вылета (понедельник недели): [week_start, n, median, p90]
    weeks: dict[str, list[int]] = {}
    for o in trips:
        d = datetime.fromisoformat(o["depart"]).date()
        weeks.setdefault((d - timedelta(days=d.weekday())).isoformat(), []).append(o["price"])
    by_week = [[w, q["n"], q["median"], q["p90"]] for w, q in sorted((w, quantiles(p)) for w, p in weeks.items())]
    direct = [o["price"] for o in trips if (o["transfers"] or 0) == 0 and (o["return_transfers"] or 0) == 0]
    le1 = [o["price"] for o in trips if (o["transfers"] or 0) <= 1 and (o["return_transfers"] or 0) <= 1]
    return {
        "status": "ok",
        "airport": c["iata"],
        "all": quantiles([o["price"] for o in trips]),
        "max1transfer": quantiles(le1),
        "direct": quantiles(direct),
        "by_week": by_week,
        "airlines": sorted({o["airline"] for o in trips if o["airline"]}),
    }


def collect_rail(s: PoliteSession, c: dict) -> dict:
    code = c.get("express")
    if c.get("iata") == "MOW" or code == tutu.MOSCOW:
        return {"status": "origin"}
    if not code:
        return {"status": "no_station"}
    there = tutu.trips(s, tutu.MOSCOW, code)
    back = tutu.trips(s, code, tutu.MOSCOW)
    if not there and not back:
        return {"status": "empty"}
    # компактно: [номер, фирменный 0/1, часы, {класс: цена}]
    pack = lambda ts: [[t["train"], int(t["firm"]), t["hours"], t["classes"]] for t in ts]  # noqa: E731
    return {"status": "ok", "station": code, "there": pack(there), "back": pack(back)}


def collect_bus(s: PoliteSession, c: dict, rail: dict | None) -> dict:
    """Автобус — только для городов РФ, куда нет поезда (замена ЖД)."""
    if c["abroad"] == "1" or c.get("iata") == "MOW":
        return {"status": "not_needed"}
    if (rail or {}).get("status") == "ok" and (rail.get("there") or rail.get("back")):
        return {"status": "not_needed"}
    res = atlas.trips(s, c["city"].split(",")[0])
    if not res["there"] and not res["back"]:
        return {"status": "empty", "source": "atlasbus"}
    return {"status": "ok", "source": "atlasbus", **res}


def pack_items(hotels: list[dict]) -> list[list]:
    # компактно для Excel: [цена/ночь RUB, звёзды, рейтинг, расстояние м, тип, апартаменты 0/1]
    return [[h["price"], h["stars"], h["rating"], h["dist_m"], h["kind"], int(h["apart"])] for h in hotels]


def collect_dated(s: PoliteSession, slug: str, windows: list[int], nights: int) -> list[dict]:
    """Отели на опорные даты заезда (сегодня + N дней) — для приблизительного расчёта на даты поездки."""
    out = []
    for offset in windows:
        checkin = (date.today() + timedelta(days=offset)).isoformat()
        try:
            res = ostrovok.fetch_dated(s, slug, checkin, nights)
        except ostrovok.Blocked as e:
            log.warning("ostrovok (даты) заблокировал: %s", e)
            continue
        except Exception as e:  # noqa: BLE001
            log.warning("ostrovok (даты) %s %s: %s", slug, checkin, e)
            continue
        if res["hotels"]:
            out.append({"checkin": checkin, "nights": nights, "n": len(res["hotels"]),
                        "price_all": quantiles([h["price"] for h in res["hotels"]]), "items": pack_items(res["hotels"])})
    return out


def collect_hotels(s: PoliteSession, c: dict, pages: int, windows: list[int], nights: int) -> tuple[dict, list[dict]]:
    slug = c.get("ostrovok_slug")
    if not slug:
        return {"status": "no_slug"}, []
    try:
        res = ostrovok.fetch_city(s, slug, pages)
    except ostrovok.Blocked as e:
        log.warning("ostrovok заблокировал: %s", e)
        return {"status": "blocked"}, []
    dated = collect_dated(s, slug, windows, nights) if windows else []
    hotels = res["hotels"]
    if not hotels:
        return {"status": "empty"}, []
    kinds = sorted({h["kind"] for h in hotels if h["kind"]})
    items = pack_items(hotels)
    summary = {
        "status": "ok",
        "source": "ostrovok",
        "slug": slug,
        "total_in_city": res["total"],
        "price_all": quantiles([h["price"] for h in hotels]),
        "kinds": kinds,
        "items": items,
        "dated": dated,
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
    ap.add_argument("--hotel-windows", default=env("HOTEL_WINDOWS") or "14",
                    help="опорные заезды для цен на даты: дней от сегодня, через запятую; пусто = не собирать")
    ap.add_argument("--hotel-nights", type=int, default=int(env("HOTEL_NIGHTS") or 7))
    ap.add_argument("--shard", type=int, default=0, help="номер потока (0..of-1) при параллельном сборе")
    ap.add_argument("--of", type=int, default=1, help="число параллельных потоков")
    ap.add_argument("--out", default=str(DATA_DIR / "prices.json"))
    ap.add_argument("--previous", default=str(DATA_DIR / "prices.json"), help="прошлый снимок (для сохранения старых данных при сбоях)")
    args = ap.parse_args()
    setup_logging()

    sources = set(args.sources.split(","))
    cities = load_cities()
    if args.cities:
        wanted = {x.strip() for x in args.cities.split(",")}
        cities = [c for c in cities if c["city"] in wanted]
    if args.limit:
        cities = cities[: args.limit]
    if args.of > 1:
        cities = [c for i, c in enumerate(cities) if i % args.of == args.shard]
    windows = [int(x) for x in args.hotel_windows.split(",") if x.strip()]

    out_path = Path(args.out)
    # Предыдущий снимок: если источник сегодня упал, оставляем прошлые данные с их датой
    prev_path = Path(args.previous)
    previous, prev_doc = {}, {}
    if prev_path.exists():
        try:
            prev_doc = json.loads(prev_path.read_text(encoding="utf-8"))
            previous = {c["id"]: c for c in prev_doc.get("cities", [])}
        except (ValueError, KeyError):
            previous, prev_doc = {}, {}

    months = aviasales.months_ahead(args.months)
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    s_air, s_rail, s_hot = PoliteSession(0.3, 0.6), PoliteSession(0.8, 1.5), PoliteSession(1.5, 3.0)
    s_bus = PoliteSession(1.0, 2.0)
    # в режиме потоков файлы отелей пишутся рядом с файлом потока — только свои города
    hotels_dir = (out_path.parent / "hotels") if args.of > 1 else (DATA_DIR / "hotels")
    hotels_dir.mkdir(parents=True, exist_ok=True)

    result = {"generated_at": now, "months": months, "hotel_windows": windows, "hotel_nights": args.hotel_nights,
              "sources": sorted(sources), "origin": {"city": "Москва", "iata": "MOW", "express": tutu.MOSCOW}}
    if "fx" in sources:
        try:
            result["fx"] = cbr.fetch_rates()
        except Exception as e:  # noqa: BLE001
            log.warning("fx: %s", e)
    if "fx" not in result and prev_doc.get("fx"):
        result["fx"] = prev_doc["fx"]  # курсы из прошлого снимка, их дата — в fx.date

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
            ("bus", lambda: collect_bus(s_bus, c, row.get("rail") or prev.get("rail"))),
            ("hotels", lambda: collect_hotels(s_hot, c, args.hotel_pages, windows, args.hotel_nights)),
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
                        (hotels_dir / f"{c['id']}.json").write_text(
                            json.dumps({"city": c["city"], "updated_at": now, "slug": data["slug"], "hotels": full}, ensure_ascii=False),
                            encoding="utf-8",
                        )
                data["updated_at"] = now
            except Exception as e:  # noqa: BLE001
                log.warning("%s %s: %s", src, c["city"], e)
                data = {"status": "error", "error": f"{type(e).__name__}: {str(e)[:200]}", "updated_at": now}
            # если сегодня неудача, а раньше были данные — сохраняем старые с пометкой
            if data["status"] not in ("ok", "not_needed") and prev.get(src, {}).get("status") == "ok":
                kept = dict(prev[src])
                kept["stale"] = True
                kept["last_attempt"] = {"at": now, "status": data["status"]}
                data = kept
            row[src] = data
        log.info("[%d/%d] %s: air=%s rail=%s bus=%s hotels=%s", n, len(cities), c["city"],
                 row.get("air", {}).get("status"), row.get("rail", {}).get("status"), row.get("bus", {}).get("status"),
                 row.get("hotels", {}).get("status"))
        rows.append(row)

    if args.of > 1:
        # поток пишет только свои города; объединяет collector.merge
        result["cities"] = rows
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(result, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        log.info("Поток %d/%d: записано %d городов в %s", args.shard, args.of, len(rows), out_path)
        return

    # города, которые не обрабатывались в этом запуске, сохраняем из прошлого снимка
    done = {r["id"] for r in rows}
    rows += [p for pid, p in previous.items() if pid not in done]
    rows.sort(key=lambda r: r["id"])
    result["cities"] = rows
    result["coverage"] = {
        src: sum(1 for r in rows if r.get(src, {}).get("status") == "ok" and not r[src].get("stale")) for src in ("air", "rail", "bus", "hotels")
    }
    result["coverage"]["total"] = len(rows)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(result, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    log.info("Записано %s: покрытие %s", out_path, result["coverage"])


if __name__ == "__main__":
    main()
