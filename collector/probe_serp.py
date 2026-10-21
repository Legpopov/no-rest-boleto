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
    t0 = time.time()
    seen = {}
    for page in range(1, 6):
        pb = dict(body, page=page); pb.pop("map_hotels")
        for attempt in range(1, 6):
            d = s.post(f"{BASE}/hotel/search/v2/site/serp?session={sid}", json=pb, headers=hdr).json()
            hs = d.get("hotels") or []
            if hs:
                break
            time.sleep(1.0)
        for h in hs:
            seen[h["master_id"]] = h
        log.info("page %d (попыток %d, %.1f c): отелей %d, накоплено %d, finished=%s", page, attempt, time.time() - t0, len(hs), len(seen), d.get("search_finished"))
        if len(hs) < 20:
            break
    h = next(iter(seen.values()))
    sv = h.get("static_vm") or {}
    log.info("static_vm ключи: %s", list(sv.keys()))
    log.info("static_vm выборка: %s", json.dumps({k: sv.get(k) for k in ("name", "kind", "star_rating", "stars", "rating", "latitude", "longitude", "is_apart_like", "hotel_kind")}, ensure_ascii=False))
    r0 = (h.get("rates") or [{}])[0]
    log.info("rate ключи: %s", list(r0.keys()))
    def find_prices(o, path=""):
        if isinstance(o, dict):
            for k, v in o.items():
                if any(w in k for w in ("amount", "price", "total", "daily")) and not isinstance(v, (dict, list)):
                    log.info("  %s.%s = %s", path, k, v)
                find_prices(v, f"{path}.{k}")
        elif isinstance(o, list) and o:
            find_prices(o[0], path + "[0]")
    find_prices(r0)


if __name__ == "__main__":
    main()
