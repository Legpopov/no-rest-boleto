// Логика API сайта — общая для Netlify Functions (netlify/functions/*.mjs)
// и Cloudflare Pages Functions (functions/api/*.js). Обёртки только передают URL запроса и токен.
//
//   flights(url, token) — GET /api/flights?origin=MOW&destination=KZN&depart=2026-11-15&return=2026-11-18
//     Aviasales Data API: варианты на даты + календарь цен ±3 дня (return можно не указывать — в одну сторону).
//     Данные — кэш поисков пользователей Aviasales за последние дни, не живой поиск.
//   trains(url)         — GET /api/trains?from=2000000&to=2060615
//     Поезда между станциями «Экспресс» (Туту через suggest.travelpayouts.com), цены «от» без привязки к дате.

const AVIA_API = 'https://api.travelpayouts.com/aviasales/v3/prices_for_dates';
const TRAINS_API = 'https://suggest.travelpayouts.com/search';
const WINDOW = 3;
const HORIZON = 365;
const IATA = /^[A-Z]{3}$/;
const DAY = /^\d{4}-\d{2}-\d{2}$/;
const CODE = /^\d{7}$/;

// кэш в памяти экземпляра функции (живёт, пока экземпляр не выгружен)
const mem = new Map();
async function memo(key, ttlMs, load) {
  const hit = mem.get(key);
  if (hit && Date.now() - hit.at < ttlMs) return hit.data;
  const data = await load();
  mem.set(key, { at: Date.now(), data });
  if (mem.size > 500) mem.delete(mem.keys().next().value);
  return data;
}

// ttl = [секунд в браузере, секунд на CDN]; null — не кэшировать (ошибки)
function json(body, status = 200, ttl = null) {
  const headers = { 'Content-Type': 'application/json; charset=utf-8' };
  if (ttl) {
    const [browser, cdn] = ttl;
    // s-maxage — для кэша Cloudflare; Netlify-CDN-Cache-Control читает только Netlify
    headers['Cache-Control'] = `public, max-age=${browser}, s-maxage=${cdn}`;
    headers['Netlify-CDN-Cache-Control'] = `public, s-maxage=${cdn}, stale-while-revalidate=${Math.round(cdn / 2)}, durable`;
  } else {
    headers['Cache-Control'] = 'no-store';
  }
  return new Response(JSON.stringify(body), { status, headers });
}

const addDays = (iso, n) => {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10);
};

async function tp(token, params) {
  const url = new URL(AVIA_API);
  for (const [k, v] of Object.entries({ sorting: 'price', currency: 'rub', unique: 'false', ...params })) url.searchParams.set(k, String(v));
  return memo(url.toString(), 30 * 60 * 1000, async () => {
    const resp = await fetch(url, { headers: { 'X-Access-Token': token, 'Accept-Encoding': 'gzip, deflate' } });
    if (resp.status === 429) throw Object.assign(new Error('Превышен лимит запросов к API, попробуйте через минуту'), { status: 429 });
    if (!resp.ok) throw Object.assign(new Error(`API ответил ${resp.status}`), { status: 502 });
    const body = await resp.json();
    if (body.success === false) throw Object.assign(new Error(body.error || 'ошибка API'), { status: 502 });
    return body.data || [];
  });
}

const norm = (o) => ({
  price: o.price,
  depart_at: o.departure_at || null,
  return_at: o.return_at || null,
  transfers: o.transfers ?? null,
  return_transfers: o.return_transfers ?? null,
  airline: o.airline || null,
  flight_number: o.flight_number || null,
  duration_to: o.duration_to ?? o.duration ?? null,
  duration_back: o.duration_back ?? null,
  origin_airport: o.origin_airport || null,
  destination_airport: o.destination_airport || null,
  gate: o.gate || null,
  link: o.link ? `https://www.aviasales.ru${o.link}` : null,
});

export async function flights(url, token) {
  const q = url.searchParams;
  const origin = (q.get('origin') || 'MOW').toUpperCase();
  const destination = (q.get('destination') || '').toUpperCase();
  const depart = q.get('depart') || '';
  const ret = q.get('return') || '';

  if (!IATA.test(origin) || !IATA.test(destination)) return json({ error: 'Укажите коды IATA откуда и куда (3 латинские буквы)' }, 400);
  if (origin === destination) return json({ error: 'Пункты отправления и назначения совпадают' }, 400);
  if (!DAY.test(depart) || (ret && !DAY.test(ret))) return json({ error: 'Даты в формате ГГГГ-ММ-ДД' }, 400);
  if (ret && ret < depart) return json({ error: 'Дата возврата раньше даты вылета' }, 400);
  const today = new Date().toISOString().slice(0, 10);
  if (depart < today) return json({ error: 'Дата вылета уже прошла' }, 400);
  // дальше года кэш Aviasales пуст; ограничение заодно сужает перебор дат, тратящий лимит токена
  if ((ret || depart) > addDays(today, HORIZON)) return json({ error: `Даты не дальше ${HORIZON} дней от сегодня` }, 400);
  if (!token) return json({ error: 'На сервере не задан TRAVELPAYOUTS_TOKEN' }, 500);

  const oneWay = !ret;
  const base = { origin, destination, one_way: String(oneWay) };

  // месяцы, которые покрывают окно ±3 дня вокруг дат вылета и возврата
  const months = (d) => [...new Set([addDays(d, -WINDOW), d, addDays(d, WINDOW)].filter((x) => x >= today).map((x) => x.slice(0, 7)))];
  const depMonths = months(depart);
  const retMonths = oneWay ? [null] : months(ret);

  try {
    const exactReq = tp(token, { ...base, departure_at: depart, ...(oneWay ? {} : { return_at: ret }), limit: 60 });
    const monthReqs = [];
    for (const dm of depMonths) for (const rm of retMonths) monthReqs.push(tp(token, { ...base, departure_at: dm, ...(rm ? { return_at: rm } : {}), limit: 1000 }));
    const [exact, ...monthly] = await Promise.all([exactReq, ...monthReqs]);

    // календарь: минимальная цена для каждой пары дат в окне ±3 дня —
    // отдельно для любых пересадок (any), не более одной (le1) и прямых (direct)
    const lo = addDays(depart, -WINDOW), hi = addDays(depart, WINDOW);
    const rlo = oneWay ? null : addDays(ret, -WINDOW), rhi = oneWay ? null : addDays(ret, WINDOW);
    const grid = new Map();
    for (const o of [...exact, ...monthly.flat()]) {
      const d = (o.departure_at || '').slice(0, 10);
      const r = oneWay ? null : (o.return_at || '').slice(0, 10);
      if (d < lo || d > hi) continue;
      if (!oneWay && (r < rlo || r > rhi || r < d)) continue;
      const key = `${d}|${r}`;
      const cell = grid.get(key) || { depart: d, return: r, any: null, le1: null, direct: null };
      const tr = Math.max(o.transfers || 0, oneWay ? 0 : o.return_transfers || 0);
      const put = (k) => { if (cell[k] == null || o.price < cell[k]) cell[k] = o.price; };
      put('any');
      if (tr <= 1) put('le1');
      if (tr === 0) put('direct');
      grid.set(key, cell);
    }

    // браузер — 10 минут, CDN — 1 час: одинаковые запросы не тратят лимит API
    return json({
      origin, destination, depart, return: ret || null, one_way: oneWay,
      offers: exact.map(norm).sort((a, b) => a.price - b.price),
      calendar: [...grid.values()].sort((a, b) => (a.depart + a.return).localeCompare(b.depart + b.return)),
      window: WINDOW,
      fetched_at: new Date().toISOString(),
      source: 'Aviasales Data API: кэш поисков за последние дни',
    }, 200, [600, 3600]);
  } catch (e) {
    return json({ error: e.message }, e.status || 502);
  }
}

export async function trains(url) {
  const q = url.searchParams;
  const from = q.get('from') || '';
  const to = q.get('to') || '';
  if (!CODE.test(from) || !CODE.test(to)) return json({ error: 'Коды станций «Экспресс» — 7 цифр' }, 400);

  const api = `${TRAINS_API}?service=tutu_trains&term=${from}&term2=${to}`;
  try {
    const data = await memo(api, 6 * 60 * 60 * 1000, async () => {
      const resp = await fetch(api);
      if (!resp.ok) throw Object.assign(new Error(`Сервис расписаний ответил ${resp.status}`), { status: 502 });
      return resp.json();
    });
    const trips = (data.trips || [])
      .map((t) => ({
        train: t.trainNumber,
        name: t.name || '',
        firm: Boolean(t.firm),
        departure: (t.departureTime || '').slice(0, 5),
        arrival: (t.arrivalTime || '').slice(0, 5),
        from_station: t.departureStation,
        to_station: t.arrivalStation,
        hours: Math.round(Number(t.travelTimeInSeconds || 0) / 360) / 10,
        classes: Object.fromEntries((t.categories || []).filter((c) => c.price).map((c) => [c.type, c.price])),
      }))
      .filter((t) => Object.keys(t.classes).length)
      .sort((a, b) => a.departure.localeCompare(b.departure));
    // браузер — 1 час, CDN — 6 часов
    return json({ from, to, trips, fetched_at: new Date().toISOString(), source: 'Туту (через Travelpayouts), цены от, без привязки к дате' }, 200, [3600, 21600]);
  } catch (e) {
    return json({ error: e.status ? e.message : `Сервис расписаний недоступен: ${e.message}` }, 502);
  }
}

// Кэш на границе Cloudflare (Cache API) для Pages Functions: ответы функций Cloudflare сам не кэширует.
// Срок берётся из s-maxage в Cache-Control ответа; ошибки (no-store) не кэшируются.
// Ключ кэша — только параметры params { имя: значение по умолчанию } в верхнем регистре:
// лишние вроде ?x=1, mow вместо MOW или явный origin=MOW вместо пропущенного
// не создают новых записей и не обходят кэш, расходуя лимит токена.
export async function edgeCached(ctx, make, params) {
  const cache = globalThis.caches?.default;
  if (!cache || ctx.request.method !== 'GET') return make();
  const url = new URL(ctx.request.url);
  const canon = new URL(url.origin + url.pathname);
  for (const [p, dflt] of Object.entries(params)) canon.searchParams.set(p, (url.searchParams.get(p) || dflt).toUpperCase());
  const key = new Request(canon.toString(), { method: 'GET' });
  try {
    const hit = await cache.match(key);
    if (hit) return hit;
  } catch { /* кэш недоступен — работаем без него */ }
  const resp = await make();
  if (resp.ok && /s-maxage=\d+/.test(resp.headers.get('Cache-Control') || '')) {
    try { ctx.waitUntil(cache.put(key, resp.clone())); } catch { /* например, локальный запуск */ }
  }
  return resp;
}
