"""Проверка: отвечает ли поиск Островка на даты без браузера (обычный POST JSON).

Запуск:  python -m collector.probe_serp 1993 2026-11-15 2026-11-18
"""
from __future__ import annotations

import json
import sys
import uuid

from .common import PoliteSession, setup_logging, log

URL = "https://ostrovok.ru/hotel/search/v2/site/serp"


def main() -> None:
    setup_logging()
    region_id, arrival, departure = int(sys.argv[1]), sys.argv[2], sys.argv[3]
    s = PoliteSession(0, 0)
    sid = str(uuid.uuid4())
    body = {
        "session_params": {"arrival_date": arrival, "departure_date": departure, "currency": "RUB", "language": "ru",
                           "paxes": [{"adults": 1}], "region_id": region_id, "search_uuid": str(uuid.uuid4())},
        "filters": {}, "session_id": sid, "map_hotels": True,
    }
    r = s.post(f"{URL}?session={sid}", json=body, headers={"Referer": "https://ostrovok.ru/", "Origin": "https://ostrovok.ru"})
    log.info("status %s, %d байт", r.status_code, len(r.content))
    try:
        d = r.json()
    except ValueError:
        print(r.text[:500])
        return
    hs = d.get("map_hotels") or []
    log.info("map_hotels: %d, ключи ответа: %s", len(hs), list(d.keys()))
    for h in hs[:5]:
        print(json.dumps({k: h.get(k) for k in ("name", "stars", "rating", "price", "latitude", "longitude")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
