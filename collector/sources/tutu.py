"""Поезда: цены по классам вагонов (Туту, через публичный сервис Travelpayouts).

Цены — минимальные по поезду для каждого класса, без привязки к дате.
Коды станций — система «Экспресс» (Москва 2000000), берутся из подсказок tutu.ru.
"""
from __future__ import annotations

from ..common import PoliteSession, log

TRIPS = "https://suggest.travelpayouts.com/search"
SUGGEST = "https://www.tutu.ru/suggest/railway_simple/"
MOSCOW = "2000000"


def _norm(s: str) -> str:
    return s.lower().replace("ё", "е").replace("-", " ").strip()


def resolve_station(session: PoliteSession, city: str) -> str | None:
    """Код «Экспресс» для города (общегородской, если есть)."""
    resp = session.get(SUGGEST, params={"name": city})
    resp.raise_for_status()
    items = resp.json() or []
    exact = [i for i in items if _norm(i.get("value", "")) == _norm(city)]
    best = (exact or [i for i in items if _norm(i.get("value", "")).startswith(_norm(city))] or [None])[0]
    return best["id"] if best else None


def trips(session: PoliteSession, from_code: str, to_code: str) -> list[dict]:
    """Поезда между станциями: [{train, name, firm, hours, classes: {coupe: 2616, ...}}]."""
    resp = session.get(TRIPS, params={"service": "tutu_trains", "term": from_code, "term2": to_code})
    resp.raise_for_status()
    body = resp.json()
    out = []
    # если поездов нет, сервис отвечает пустым списком вместо объекта
    for t in (body.get("trips") if isinstance(body, dict) else None) or []:
        classes = {c["type"]: c["price"] for c in t.get("categories") or [] if c.get("price")}
        if not classes:
            continue
        out.append(
            {
                "train": t.get("trainNumber"),
                "name": t.get("name") or "",
                "firm": bool(t.get("firm")),
                "hours": round(int(t.get("travelTimeInSeconds") or 0) / 3600, 1),
                "departure": t.get("departureTime"),
                "classes": classes,
            }
        )
    if not out:
        log.info("tutu: нет поездов %s -> %s", from_code, to_code)
    return out
