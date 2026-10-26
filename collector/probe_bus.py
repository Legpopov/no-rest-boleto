"""Проверка автобусного поиска Туту (offers-api) без браузера.

Запуск:  python -m collector.probe_bus "Донецк,Луганск,Брянск" 2026-11-10,2026-11-17
"""
from __future__ import annotations

import json
import sys
import uuid

from .common import PoliteSession, log, setup_logging

GEO = "https://bus.tutu.ru/bus/v1/geo/geopoint_incline/"
OFFERS = "https://offers-api.tutu.ru/mtp/offers"


def geo_id(s: PoliteSession, name: str) -> int | None:
    r = s.get(GEO, params={"name": name})
    r.raise_for_status()
    return ((r.json().get("data") or {}).get("geopoint") or {}).get("id")


def main() -> None:
    setup_logging()
    cities = sys.argv[1].split(",")
    dates = sys.argv[2].split(",")
    s = PoliteSession(0.5, 1.0)
    s.headers.update({"Origin": "https://bus.tutu.ru", "Referer": "https://bus.tutu.ru/"})
    moscow = geo_id(s, "Москва")
    for city in cities:
        to = geo_id(s, city)
        for day in dates:
            sid = str(uuid.uuid4())
            body = {"routes": [{"departureId": str(moscow), "arrivalId": str(to), "departureDate": day, "geoType": "BUS"}],
                    "passengers": {"adultCount": 1, "children": []}, "userData": {"sessionId": sid}, "sessionId": sid,
                    "searchId": str(uuid.uuid4()), "source": "offers"}
            import time
            voyages, txt, blocks = 0, "", []
            for attempt in range(1, 9):
                r = s.post(OFFERS, json=body)
                txt = r.text
                try:
                    data = r.json()
                except ValueError:
                    log.info("%s %s: %s не JSON: %s", city, day, r.status_code, txt[:200])
                    break
                blocks = data if isinstance(data, list) else [data]
                voyages = sum(len(((b.get("dictionary") or {}).get("bus") or {}).get("voyages") or {}) for b in blocks)
                keys = sorted({k for b in blocks for k in b.keys()})
                log.info("  попытка %d: рейсов %d, ключи %s, finished=%s", attempt, voyages, keys, [b.get("isFinished", b.get("finished", b.get("status"))) for b in blocks])
                if voyages:
                    break
                time.sleep(2)
            prices = sorted({int(float(m)) for m in __import__("re").findall(r'"(?:amount|total|value)":\s*"?(\d{3,6}(?:\.\d+)?)', txt)})[:12]
            log.info("%s (%s) %s: status %s, частей ответа %d, рейсов %d, похожие цены %s", city, to, day, r.status_code, len(blocks), voyages, prices)
            if voyages and city == cities[0]:
                b = next(b for b in blocks if ((b.get("dictionary") or {}).get("bus") or {}).get("voyages"))
                print(json.dumps({k: (list(v.keys())[:5] if isinstance(v, dict) else v) for k, v in b.items()}, ensure_ascii=False)[:1500])
                print(txt[:3000])


if __name__ == "__main__":
    main()
