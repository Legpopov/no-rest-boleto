"""Официальные курсы валют ЦБ РФ (XML_daily.asp)."""
from __future__ import annotations

import xml.etree.ElementTree as ET

from ..common import PoliteSession

URL = "https://www.cbr.ru/scripts/XML_daily.asp"


def fetch_rates(session: PoliteSession | None = None) -> dict:
    """Возвращает {'date': 'ДД.ММ.ГГГГ', 'rates': {'USD': 81.23, ...}} — рублей за 1 единицу валюты."""
    session = session or PoliteSession(0, 0)
    resp = session.get(URL)
    resp.raise_for_status()
    root = ET.fromstring(resp.content)
    rates = {}
    for v in root.findall("Valute"):
        code = v.findtext("CharCode")
        nominal = int(v.findtext("Nominal"))
        value = float(v.findtext("Value").replace(",", "."))
        rates[code] = round(value / nominal, 6)
    return {"date": root.get("Date"), "rates": rates}
