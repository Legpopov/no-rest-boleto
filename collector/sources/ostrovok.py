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
    hotels = resp.json().get("hotels", [])
    regions = {}
    for h in hotels:
        if h.get("slug"):
            regions.setdefault(h["slug"], h)
    cands = list(regions.values())
    exact = [h for h in cands if _norm(h.get("region_name", "")) == _norm(city) and (not country_code or h.get("country_code") == country_code)]
    same_country = [h for h in cands if not country_code or h.get("country_code") == country_code]
    best = (exact or same_country or [None])[0]
    if not best:
        return None
    return {"slug": best["slug"], "region_id": best.get("region_id"), "region_name": best.get("region_name"), "exact": bool(exact)}


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
