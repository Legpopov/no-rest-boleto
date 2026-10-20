"""Проверка доступности сайтов с отелями с текущего IP.

Ничего не парсит: только смотрит код ответа, размер и признаки страницы-проверки на бота.
Запуск:  python -m collector.probe_hotels
"""
from __future__ import annotations

import json

from .common import PoliteSession, setup_logging, log

TARGETS = {
    "ostrovok": "https://ostrovok.ru/hotel/russia/kazan/",
    "yandex_travel": "https://travel.yandex.ru/hotels/kazan/",
    "sutochno": "https://sutochno.ru/",
    "booking": "https://www.booking.com/city/ru/kazan.ru.html",
}
BOT_MARKERS = ("captcha", "smartcaptcha", "showcaptcha", "challenge", "are you a robot", "access denied")


def main() -> None:
    setup_logging()
    s = PoliteSession(1, 2)
    try:
        ip = s.get("https://api.country.is/", timeout=10).json()
    except Exception as e:  # noqa: BLE001
        ip = {"error": str(e)}
    log.info("Выход в интернет: %s", ip)
    report = {"egress": ip, "targets": {}}
    for name, url in TARGETS.items():
        try:
            r = s.get(url, timeout=25, allow_redirects=True)
            low = r.text[:20000].lower()
            marker = next((m for m in BOT_MARKERS if m in low), None)
            res = {"status": r.status_code, "bytes": len(r.content), "final_url": r.url, "bot_marker": marker}
        except Exception as e:  # noqa: BLE001
            res = {"error": f"{type(e).__name__}: {str(e)[:120]}"}
        report["targets"][name] = res
        log.info("%-14s %s", name, res)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
