"""Проверка: отвечает ли поиск Островка на даты без браузера.

Как браузер: сначала GET страницы города с датами (cookies + id поиска из __NEXT_DATA__),
затем POST /hotel/search/v2/site/serp с опросом до search_finished.
Запуск:  python -m collector.probe_serp russia/kazan 2026-11-15 2026-11-18
"""
from __future__ import annotations

import json
import re
import sys
import time
import uuid

from .common import PoliteSession, setup_logging, log

BASE = "https://ostrovok.ru"


def main() -> None:
    setup_logging()
    slug, arrival, departure = sys.argv[1], sys.argv[2], sys.argv[3]
    s = PoliteSession(0, 0)
    d1 = "{2}.{1}.{0}".format(*arrival.split("-"))
    d2 = "{2}.{1}.{0}".format(*departure.split("-"))
    page_url = f"{BASE}/hotel/{slug}/?dates={d1}-{d2}&guests=1"
    page = s.get(page_url)
    log.info("страница: %s, cookies: %s", page.status_code, list(s.cookies.keys()))
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', page.text, re.S)
    props = json.loads(m.group(1))["props"]["pageProps"]
    srd = props.get("searchRegionData") or {}
    log.info("ключи searchRegionData: %s", list(srd.keys())[:30])
    region_id = srd.get("id") or srd.get("regionId") or srd.get("region_id") or (int(sys.argv[4]) if len(sys.argv) > 4 else None)
    search_uuid = props.get("searchUuid") or str(uuid.uuid4())
    log.info("region_id=%s searchUuid=%s", region_id, search_uuid)
    sid = str(uuid.uuid4())
    body = {
        "session_params": {"arrival_date": arrival, "departure_date": departure, "currency": "RUB", "language": "ru",
                           "paxes": [{"adults": 1}], "region_id": region_id, "search_uuid": search_uuid},
        "filters": {}, "session_id": sid, "map_hotels": True,
    }
    hdr = {"Referer": page_url, "Origin": BASE, "X-CSRFToken": s.cookies.get("csrftoken", "")}
    d = {}
    for attempt in range(1, 21):
        r = s.post(f"{BASE}/hotel/search/v2/site/serp?session={sid}", json=body, headers=hdr)
        d = r.json()
        log.info("попытка %d: finished=%s total_av=%s map=%d", attempt, d.get("search_finished"), d.get("total_av_hotels"),
                 len(d.get("map_hotels") or []))
        if d.get("search_finished"):
            break
        time.sleep(1.5)
    hs = d.get("map_hotels") or []
    prices = sorted(h["price"] for h in hs if h.get("price"))
    log.info("итог: %d отелей с ценой, медиана за весь срок %s", len(prices), prices[len(prices) // 2] if prices else None)
    for h in hs[:3]:
        print(json.dumps({k: h.get(k) for k in ("name", "stars", "rating", "price", "latitude", "longitude")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
