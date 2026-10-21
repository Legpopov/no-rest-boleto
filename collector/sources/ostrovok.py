"""Отели с публичной выдачи ostrovok.ru.

Страница города (без дат) содержит JSON __NEXT_DATA__ с 20 отелями на страницу:
цена «от» за ночь, звёзды, рейтинг, расстояние до центра, координаты.
При странице-проверке на бота источник пропускается — обход не выполняется.
"""
from __future__ import annotations

import json
import re

from ..common import PoliteSession, log

BASE = "https://ostrovok.ru"
NEXT_DATA = re.compile(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', re.S)


class Blocked(RuntimeError):
    """Сайт показал проверку на бота."""


def _norm(s: str) -> str:
    return s.lower().replace("ё", "е").replace("-", " ").strip()


def resolve_slug(session: PoliteSession, city: str, country_code: str) -> dict | None:
    """Адрес города в каталоге Островка, например {'slug': 'australia/sydney', 'region_id': 3341}."""
    resp = session.get(f"{BASE}/api/site/multicomplete.json", params={"query": city, "locale": "ru"})
    resp.raise_for_status()
    body = resp.json()
    # регионы: {"type": "City", "name": "Сидней, Новый Южный Уэльс", "country_code": "AU", "slug": ...}
    cands = [
        {"slug": r["slug"], "region_id": r.get("id"), "region_name": r.get("name", ""), "country_code": r.get("country_code"),
         "currency": r.get("currency_code"), "type": r.get("type")}
        for r in body.get("regions") or [] if r.get("slug")
    ]
    if not cands:  # запасной путь: регион первого найденного отеля
        cands = [
            {"slug": h["slug"], "region_id": h.get("region_id"), "region_name": h.get("region_name", ""), "country_code": h.get("country_code"),
             "currency": None, "type": "City"}
            for h in body.get("hotels") or [] if h.get("slug")
        ]
    in_country = [c for c in cands if not country_code or c["country_code"] == country_code]
    exact = [c for c in in_country if _norm(c["region_name"].split(",")[0]) == _norm(city)]
    cities = [c for c in in_country if c["type"] == "City"]
    best = (exact or cities or [None])[0]
    if not best:
        return None
    best["exact"] = bool(exact)
    return best


def fetch_page(session: PoliteSession, slug: str, page: int = 1) -> tuple[list[dict], dict]:
    url = f"{BASE}/hotel/{slug}/" + (f"?page={page}" if page > 1 else "")
    resp = session.get(url)
    low = resp.text[:20000].lower()
    if "captcha" in low or "challenge" in resp.url:
        raise Blocked(url)
    if resp.status_code == 404:
        return [], {}
    resp.raise_for_status()
    m = NEXT_DATA.search(resp.text)
    if not m:
        raise Blocked(f"нет __NEXT_DATA__: {url}")
    props = json.loads(m.group(1))["props"]["pageProps"]
    serp = props.get("serpData") or {}
    meta = {"total": serp.get("totalHotels"), "center": (props.get("initialPageState") or {}).get("regionCenter")}
    hotels = []
    for h in serp.get("hotels") or []:
        rate = h.get("datefullRate") or h.get("datelessRate") or {}
        loc = h.get("location") or {}
        if not rate.get("amount"):
            continue
        hotels.append(
            {
                "id": h.get("masterId"),
                "slug_id": h.get("id"),
                "name": h.get("name"),
                "kind": h.get("kind"),
                "apart": bool(h.get("isApartLike")),
                "stars": h.get("stars") or 0,
                "rating": h.get("rating"),
                "reviews": h.get("numReviews") or 0,
                "dist_m": h.get("searchRegionCenterDistance"),
                "lat": loc.get("latitude"),
                "lon": loc.get("longitude"),
                "address": loc.get("address"),
                "price": round(rate["amount"]),
                "currency": rate.get("currency", "RUB"),
            }
        )
    return hotels, meta


APART_KINDS = {"apartment", "apart-hotel", "apart_hotel", "flat", "villa", "cottage", "guesthouse", "hostel", "bnb"}


def fetch_dated(session: PoliteSession, slug: str, checkin: str, nights: int, pages: int = 5) -> dict:
    """Отели на конкретные даты: как браузер — страница с датами (cookies, id региона и поиска),
    затем постранично POST /hotel/search/v2/site/serp. Цена — средняя за ночь у самого дешёвого тарифа.
    """
    import datetime as dt
    import uuid

    d_in = dt.date.fromisoformat(checkin)
    d_out = d_in + dt.timedelta(days=nights)
    page_url = f"{BASE}/hotel/{slug}/?dates={d_in:%d.%m.%Y}-{d_out:%d.%m.%Y}&guests=1"
    resp = session.get(page_url)
    if "captcha" in resp.text[:20000].lower():
        raise Blocked(page_url)
    resp.raise_for_status()
    m = NEXT_DATA.search(resp.text)
    if not m:
        raise Blocked(f"нет __NEXT_DATA__: {page_url}")
    props = json.loads(m.group(1))["props"]["pageProps"]
    region_id = (props.get("searchRegionData") or {}).get("regionId")
    if not region_id:
        return {"checkin": checkin, "nights": nights, "hotels": []}
    sid = str(uuid.uuid4())
    body = {
        "session_params": {"arrival_date": d_in.isoformat(), "departure_date": d_out.isoformat(), "currency": "RUB",
                           "language": "ru", "paxes": [{"adults": 1}], "region_id": region_id,
                           "search_uuid": props.get("searchUuid") or str(uuid.uuid4())},
        "filters": {}, "session_id": sid,
    }
    hdr = {"Referer": page_url, "Origin": BASE, "X-CSRFToken": session.cookies.get("csrftoken", "")}
    seen: dict = {}
    for page in range(1, pages + 1):
        hs = []
        for _ in range(5):  # поиск асинхронный: первая страница может прийти пустой
            r = session.post(f"{BASE}/hotel/search/v2/site/serp?session={sid}", json=dict(body, page=page), headers=hdr)
            r.raise_for_status()
            hs = r.json().get("hotels") or []
            if hs:
                break
            import time
            time.sleep(1.0)
        for h in hs:
            sv = h.get("static_vm") or {}
            per_day = []
            for rate in h.get("rates") or []:
                for pt in (rate.get("payment_options") or {}).get("payment_types") or []:
                    v = pt.get("average_show_price_per_day")
                    if v:
                        per_day.append(float(v))
            if not per_day:
                continue
            kind = (sv.get("kind") or "").lower()
            seen[h.get("master_id")] = {
                "id": h.get("master_id"),
                "name": sv.get("name"),
                "kind": sv.get("kind"),
                "apart": kind in APART_KINDS,
                "stars": round((sv.get("star_rating") or 0) / 10),
                "rating": (sv.get("rating") or {}).get("total"),
                "dist_m": sv.get("search_region_center_distance"),
                "lat": sv.get("latitude"),
                "lon": sv.get("longitude"),
                "price": round(min(per_day)),
            }
        if len(hs) < 20:
            break
    log.info("ostrovok %s %s+%d: %d отелей на даты", slug, checkin, nights, len(seen))
    return {"checkin": checkin, "nights": nights, "hotels": list(seen.values())}


def fetch_city(session: PoliteSession, slug: str, pages: int = 5) -> dict:
    """Отели города с нескольких страниц выдачи. Дубликаты по id отбрасываются."""
    seen, out, meta = set(), [], {}
    for page in range(1, pages + 1):
        hotels, m = fetch_page(session, slug, page)
        meta = meta or m
        new = [h for h in hotels if h["id"] not in seen]
        seen.update(h["id"] for h in new)
        out.extend(new)
        if len(hotels) < 20:
            break
    log.info("ostrovok %s: %d отелей (всего в городе %s)", slug, len(out), meta.get("total"))
    return {"hotels": out, "total": meta.get("total"), "center": meta.get("center")}
