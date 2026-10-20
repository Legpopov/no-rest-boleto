"""Общие утилиты сборщика: HTTP-сессия с паузами, пути, логирование."""
from __future__ import annotations

import logging
import os
import random
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
CACHE_DIR = Path(__file__).resolve().parent / ".cache"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"
)

log = logging.getLogger("collector")


def setup_logging(level: str = "INFO") -> None:
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


class PoliteSession(requests.Session):
    """Сессия, которая делает паузу между запросами к одному хосту."""

    def __init__(self, min_delay: float = 1.0, max_delay: float = 2.5) -> None:
        super().__init__()
        self.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "ru-RU,ru;q=0.9"})
        self.min_delay, self.max_delay = min_delay, max_delay
        self._last: dict[str, float] = {}

    def request(self, method, url, *args, **kwargs):
        host = requests.utils.urlparse(url).netloc
        wait = random.uniform(self.min_delay, self.max_delay) - (time.monotonic() - self._last.get(host, 0))
        if wait > 0:
            time.sleep(wait)
        kwargs.setdefault("timeout", 30)
        try:
            return super().request(method, url, *args, **kwargs)
        finally:
            self._last[host] = time.monotonic()


def load_dotenv(path: Path = ROOT / ".env") -> None:
    """Подхватывает переменные из .env (для локального запуска; в Actions их дают секреты)."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


load_dotenv()


def env(name: str, required: bool = False) -> str | None:
    value = os.environ.get(name)
    if required and not value:
        raise RuntimeError(f"Не задана переменная окружения {name}")
    return value
