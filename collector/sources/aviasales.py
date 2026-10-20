"""Цены авиабилетов из Aviasales Data API (Travelpayouts).

Данные — кэш поисков пользователей Aviasales за последние дни, а не живой поиск.
Токен: переменная окружения TRAVELPAYOUTS_TOKEN.
Документация: https://support.travelpayouts.com/hc/en-us/articles/203956163
"""
from __future__ import annotations

from datetime import date

from ..common import PoliteSession, env, log

URL = "https://api.travelpayouts.com/aviasales/v3/prices_for_dates"


def fetch_round_trips(session: PoliteSession, origin: str, destination: str, months: list[str]) -> list[dict]:
    """Все найденные варианты туда-обратно с вылетом в указанные месяцы (YYYY-MM).

    Возвращает список {price, depart, return, transfers, return_transfers, airline, duration_min, link}.
    """
    token = env("TRAVELPAYOUTS_TOKEN", required=True)
    offers: list[dict] = []
    for month in months:
        resp = session.get(
            URL,
            params={
                "origin": origin,
                "destination": destination,
                "departure_at": month,
                "one_way": "false",
                "sorting": "price",
                "currency": "rub",
                "limit": 1000,
                "unique": "false",
            },
            headers={"X-Access-Token": token, "Accept-Encoding": "gzip, deflate"},
        )
        if resp.status_code == 429:
            log.warning("aviasales: превышен лимит запросов на %s-%s %s", origin, destination, month)
            continue
        resp.raise_for_status()
        body = resp.json()
        if not body.get("success", True):
            log.warning("aviasales: %s-%s %s: %s", origin, destination, month, body.get("error"))
            continue
        for o in body.get("data", []):
            offers.append(
                {
                    "price": o["price"],
                    "depart": (o.get("departure_at") or "")[:10],
                    "return": (o.get("return_at") or "")[:10],
                    "transfers": o.get("transfers"),
                    "return_transfers": o.get("return_transfers"),
                    "airline": o.get("airline"),
                    "duration_min": o.get("duration"),
                    "link": o.get("link"),
                }
            )
    return offers


def months_ahead(n: int, start: date | None = None) -> list[str]:
    d = start or date.today()
    out = []
    y, m = d.year, d.month
    for _ in range(n):
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out
