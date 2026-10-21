"""Объединяет результаты параллельных потоков сбора с прошлым снимком -> data/prices.json.

Запуск:  python -m collector.merge data/parts/*.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from .common import DATA_DIR, log, setup_logging


def main() -> None:
    setup_logging()
    out = DATA_DIR / "prices.json"
    base = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {"cities": []}
    cities = {c["id"]: c for c in base.get("cities", [])}
    parts = [json.loads(Path(p).read_text(encoding="utf-8")) for p in sys.argv[1:]]
    if not parts:
        raise SystemExit("нет файлов потоков")
    for p in parts:
        for c in p["cities"]:
            cities[c["id"]] = c
    head = parts[0]
    result = {k: v for k, v in head.items() if k != "cities"}
    result["generated_at"] = max(p["generated_at"] for p in parts)
    if "fx" not in result and base.get("fx"):
        result["fx"] = base["fx"]
    rows = sorted(cities.values(), key=lambda r: r["id"])
    result["cities"] = rows
    result["coverage"] = {
        src: sum(1 for r in rows if r.get(src, {}).get("status") == "ok" and not r[src].get("stale")) for src in ("air", "rail", "hotels")
    }
    result["coverage"]["total"] = len(rows)
    result["coverage"]["hotels_dated"] = sum(1 for r in rows if (r.get("hotels") or {}).get("dated"))
    result["coverage"]["air_by_week"] = sum(1 for r in rows if (r.get("air") or {}).get("by_week"))
    out.write_text(json.dumps(result, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    log.info("Объединено потоков: %d, городов: %d, покрытие %s", len(parts), len(rows), result["coverage"])


if __name__ == "__main__":
    main()
