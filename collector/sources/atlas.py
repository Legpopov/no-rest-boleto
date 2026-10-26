"""Автобусы: цены с публичных страниц маршрутов atlasbus.ru (Атлас).

Страница «/Маршруты/<откуда>/<куда>» содержит __NEXT_DATA__:
  pageProps.seoDirection.minPrices — минимальная и максимальная цена на направлении, rideCount;
  initialState.searchRides.calendar — минимальная цена по датам на ~2 недели вперёд.
Цены — в одну сторону, за 1 пассажира. При проверке на бота источник пропускается.
"""
from __future__ import annotations

import json
import re
from urllib.parse import quote

from ..common import PoliteSession, log

BASE = "https://atlasbus.ru"
NEXT_DATA = re.compile(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', re.S)


def _price(mp: dict | None) -> tuple[float | None, float | None]:
    if not mp:
        return None, None
    return mp.get("price"), mp.get("maxPrice")


def direction(session: PoliteSession, from_city: str, to_city: str) -> dict | None:
    """Цены на направлении или None, если страницы маршрута нет."""
    url = f"{BASE}/{quote('Маршруты')}/{quote(from_city)}/{quote(to_city)}"
    resp = session.get(url)
    if resp.status_code == 404:
        return None
    if resp.status_code == 403:
        raise RuntimeError("atlasbus: доступ закрыт (403)")
    resp.raise_for_status()
    m = NEXT_DATA.search(resp.text)
    if not m:
        return None
    props = json.loads(m.group(1)).get("props") or {}
    seo = (props.get("pageProps") or {}).get("seoDirection") or {}
    mins = seo.get("minPrices") or []
    low, high = _price(mins[0] if mins else None)
    cal = (((props.get("initialState") or {}).get("searchRides") or {}).get("calendar")) or []
    days = []
    for d in cal:
        p, _ = _price((d.get("minPrices") or [None])[0])
        if p and (d.get("rideCount") or 0) > 0:
            days.append([d.get("date"), round(p)])
    if not low and not days:
        return None
    return {"url": url, "low": low, "high": high, "rides": seo.get("rideCount"), "days": days}


def trips(session: PoliteSession, city: str, origin: str = "Москва") -> dict:
    """Туда и обратно. Формат для prices.json: there/back = [[метка, часов (None), цена], ...]."""
    out = {}
    for key, a, b in (("there", origin, city), ("back", city, origin)):
        d = direction(session, a, b)
        rows = []
        if d:
            rows = [[f"мин. цена на {day}", None, price] for day, price in d["days"]]
            if not rows and d["low"]:
                rows = [["мин. цена на направлении", None, round(d["low"])]]
            if d["high"] and d["high"] != d["low"]:
                rows.append(["макс. цена на направлении", None, round(d["high"])])
            out[f"{key}_url"] = d["url"]
            out[f"{key}_rides"] = d["rides"]
        out[key] = rows
    log.info("atlas %s: туда %d, обратно %d цен", city, len(out["there"]), len(out["back"]))
    return out
