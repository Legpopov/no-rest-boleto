"""Pure planning calculations. Missing evidence is never a numeric zero."""
from __future__ import annotations

from dataclasses import dataclass
import math
from statistics import mean, median


@dataclass(frozen=True)
class Calendar:
    autumn_start: int = 9
    autumn_months: int = 5
    spring_start: int = 2
    spring_months: int = 5

    def __post_init__(self):
        if not all(1 <= m <= 12 for m in (self.autumn_start, self.spring_start,
                                          self.autumn_months, self.spring_months)):
            raise ValueError("Calendar months and durations must be in 1..12")


def semester_months(launch_year, cohort, semester, calendar=Calendar()):
    if any(not isinstance(v, int) or isinstance(v, bool) or v < 1
           for v in (launch_year, cohort, semester)):
        raise ValueError("Year, cohort and semester must be positive integers")
    academic_year = launch_year + cohort - 1 + (semester - 1) // 2
    autumn = semester % 2 == 1
    start = calendar.autumn_start if autumn else calendar.spring_start
    year = academic_year + int(not autumn and start < calendar.autumn_start)
    count = calendar.autumn_months if autumn else calendar.spring_months
    return [f"{year + (start - 1 + i) // 12:04d}-{(start - 1 + i) % 12 + 1:02d}"
            for i in range(count)]


def positive(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value) and value > 0


def annual_index(current, previous):
    """Callers must establish month/class/currency/booking-horizon comparability."""
    if not positive(current) or not positive(previous):
        return None
    return current / previous - 1


def select_index(route=None, peers=(), official=None):
    """Each candidate carries a source and a decimal annual rate, e.g. 0.05."""
    def valid(item):
        return (isinstance(item, dict) and bool(item.get("source"))
                and isinstance(item.get("rate"), (int, float)) and not isinstance(item["rate"], bool)
                and math.isfinite(item["rate"]) and item["rate"] > -1)
    if valid(route):
        return {**route, "method": "route"}
    valid_peers = [p for p in peers if valid(p)]
    if valid_peers:
        return {"rate": median(p["rate"] for p in valid_peers), "method": "peers",
                "source": [p["source"] for p in valid_peers], "routes": len(valid_peers)}
    if valid(official):
        return {**official, "method": "official"}
    return None


def forecast_month(base_price, base_year, target_month, indices):
    """Compound each year's sourced index; no implicit zero growth or extrapolation."""
    year, month = map(int, target_month.split("-"))
    if not 1 <= month <= 12 or year < base_year:
        raise ValueError("Target month precedes base year or is invalid")
    if not positive(base_price):
        return {"status": "missing_base_price", "price": None, "month": target_month}
    price, used = base_price, []
    for y in range(base_year + 1, year + 1):
        index = indices.get(y)
        if not index or select_index(route=index) is None:
            return {"status": "missing_index", "price": None, "month": target_month, "year": y}
        price *= 1 + index["rate"]
        if not positive(price):
            return {"status": "invalid_forecast", "price": None, "month": target_month}
        used.append({"year": y, **index})
    return {"status": "ok", "month": target_month, "price": round(price, 2),
            "base_price": base_price, "base_year": base_year, "indices": used}


def forecast_semester(months, bases, base_year, indices):
    """Equal month weights; bases[month_number], indices[month_number][year]."""
    if not months or len(set(months)) != len(months):
        raise ValueError("Supply distinct semester months")
    results = [forecast_month(bases.get(int(m[-2:])), base_year, m,
                              indices.get(int(m[-2:]), {})) for m in months]
    complete = all(r["status"] == "ok" for r in results)
    return {"status": "ok" if complete else "insufficient_data",
            "price": round(mean(r["price"] for r in results), 2) if complete else None,
            "months": results, "covered_months": sum(r["status"] == "ok" for r in results),
            "required_months": len(results), "weighting": "equal_months"}


def choose_transport(estimates, preferences):
    if len(set(preferences)) != len(preferences):
        raise ValueError("Transport preferences must be distinct")
    for mode in preferences:
        estimate = estimates.get(mode, {})
        if estimate.get("status") == "ok" and positive(estimate.get("price")):
            return {"status": "ok", "mode": mode, **estimate}
    return {"status": "insufficient_data", "price": None, "mode": None}


def route_total(legs):
    """Explicit one-way legs only. Round-trip fares cannot price individual segments."""
    if not legs:
        return {"status": "empty_route", "price": None}
    for i, leg in enumerate(legs):
        if leg.get("basis") != "one_way":
            raise ValueError("Route legs require one-way fares, never round-trip / 2")
        if not leg.get("origin") or not leg.get("destination") or leg["origin"] == leg["destination"]:
            raise ValueError("A leg requires two distinct endpoints")
        if i and legs[i-1]["destination"] != leg["origin"]:
            raise ValueError("Route legs are not connected")
    if not all(l.get("status") == "ok" and positive(l.get("price")) for l in legs):
        return {"status": "insufficient_data", "price": None}
    return {"status": "ok", "price": round(sum(l["price"] for l in legs), 2)}
