"""Поезда: расписание и цены из API Яндекс Расписаний.

Ключ: переменная окружения YANDEX_RASP_KEY.
Документация: https://yandex.ru/dev/rasp/doc/ru/
Цены (tickets_info) приходят не для всех поездов — только где они известны Яндексу.
"""
from __future__ import annotations

from ..common import PoliteSession, env, log

BASE = "https://api.rasp.yandex.net/v3.0"
MOSCOW = "c213"


def _get(session: PoliteSession, path: str, **params) -> dict:
    params["apikey"] = env("YANDEX_RASP_KEY", required=True)
    params.setdefault("format", "json")
    params.setdefault("lang", "ru_RU")
    resp = session.get(f"{BASE}/{path}/", params=params)
    resp.raise_for_status()
    return resp.json()


def nearest_settlement(session: PoliteSession, lat: float, lon: float, distance_km: int = 30) -> dict | None:
    """Код населённого пункта Яндекса (вида c213) по координатам."""
    data = _get(session, "nearest_settlement", lat=lat, lng=lon, distance=distance_km)
    return {"code": data["code"], "title": data.get("title")} if data.get("code") else None


def search_trains(session: PoliteSession, to_code: str, day: str, from_code: str = MOSCOW) -> list[dict]:
    """Поезда на дату (YYYY-MM-DD). Возвращает список рейсов с ценами по классам вагонов."""
    data = _get(session, "search", **{"from": from_code, "to": to_code, "date": day, "transport_types": "train", "limit": 200})
    out = []
    for seg in data.get("segments", []):
        thread = seg.get("thread", {})
        places = (seg.get("tickets_info") or {}).get("places") or []
        out.append(
            {
                "number": thread.get("number"),
                "title": thread.get("title"),
                "express": thread.get("express_type"),
                "departure": seg.get("departure"),
                "arrival": seg.get("arrival"),
                "duration_min": round((seg.get("duration") or 0) / 60),
                "prices": [
                    {"class": p.get("name"), "price": p["price"]["whole"] + p["price"].get("cents", 0) / 100}
                    for p in places
                    if p.get("price")
                ],
            }
        )
    if not out:
        log.info("rasp: нет поездов %s -> %s на %s", from_code, to_code, day)
    return out
