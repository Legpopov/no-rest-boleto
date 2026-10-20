# Сайт и функции Netlify

```
netlify.toml              publish = site, functions = netlify/functions, сборка = node netlify/build-data.mjs
netlify/build-data.mjs    копирует data/ в site/data/ и строит site/data/cities.json из collector/*.csv
netlify/functions/
  flights.mjs             /api/flights — Aviasales Data API: варианты на даты + календарь ±3 дня
  trains.mjs              /api/trains  — поезда Туту (через Travelpayouts), цены по классам без даты
netlify/dev-server.mjs    локальная проверка без netlify-cli
site/                     статический фронтенд (vanilla JS + Leaflet/OpenStreetMap)
site/sample/              демо-данные: показываются, пока в site/data/ нет prices.json
```

Ответы функций кэшируются на CDN Netlify (`Netlify-CDN-Cache-Control: durable`): авиа — 1 час,
поезда — 6 часов. Одинаковые запросы разных посетителей не расходуют лимит API.
Каждый ночной коммит данных в `data/` запускает пересборку сайта.

## Локальная проверка

```bash
node netlify/build-data.mjs
```

```bash
node netlify/dev-server.mjs 8888
```

Открыть http://localhost:8888. Токен берётся из `.env` в корне репозитория (`TRAVELPAYOUTS_TOKEN=...`).
Папка `site/data/` создаётся сборкой и в git не попадает.

## Первый деплой (один раз, через сайт Netlify)

1. Зайти на https://app.netlify.com и войти через GitHub (кнопка «Sign up / Log in with GitHub»).
2. **Add new site → Import an existing project → GitHub**, разрешить доступ к репозиторию
   `Legpopov/no-rest-boleto` и выбрать его.
3. Настройки сборки подтянутся из `netlify.toml` (Branch: `main`, Build command: `node netlify/build-data.mjs`,
   Publish directory: `site`, Functions directory: `netlify/functions`) — ничего не менять.
4. До нажатия **Deploy** раскрыть **Add environment variables** (или позже:
   **Site configuration → Environment variables → Add a variable**):
   - Key: `TRAVELPAYOUTS_TOKEN`
   - Value: токен Travelpayouts (тот же, что в секретах GitHub Actions)
   - Scopes: оставить по умолчанию (как минимум **Functions**).
5. Нажать **Deploy**. Через 1–2 минуты сайт откроется по адресу вида `https://<случайное-имя>.netlify.app`.
6. **Site configuration → Site details → Change site name** → `no-rest-boleto`.
   Адрес станет https://no-rest-boleto.netlify.app (если имя свободно).
7. Если переменную добавили после первого деплоя: **Deploys → Trigger deploy → Deploy site**,
   чтобы функции получили токен.

Проверка: открыть `https://no-rest-boleto.netlify.app/api/trains?from=2000000&to=2060615` — должен
прийти JSON с поездами; на вкладке «Авиа» поиск должен находить варианты.

Альтернатива через терминал (нужен Node): `npx netlify-cli login`, затем `npx netlify-cli init`
в корне репозитория и `npx netlify-cli env:set TRAVELPAYOUTS_TOKEN <токен>`.
