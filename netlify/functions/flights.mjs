// Поиск авиабилетов по датам + цены соседних дат (±3 дня) для подсказки гибких дат.
// GET /api/flights?origin=MOW&destination=KZN&depart=2026-11-15&return=2026-11-18
//   return можно не указывать — тогда поиск в одну сторону.
// Источник: Aviasales Data API (кэш поисков пользователей за последние дни, не живой поиск).
// Токен — только из переменной окружения TRAVELPAYOUTS_TOKEN.

const API = 'https://api.travelpayouts.com/aviasales/v3/prices_for_dates';
const WINDOW = 3;
const TTL_MS = 30 * 60 * 1000;
const mem = new Map(); // кэш на время жизни экземпляра функции

const IATA = /^[A-Z]{3}$/;
const DAY = /^\d{4}-\d{2}-\d{2}$/;

const addDays = (iso, n) => {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10);
};

function json(body, status = 200, cache = true) {
  const headers = { 'Content-Type': 'application/json; charset=utf-8' };
  if (cache) {
    headers['Cache-Control'] = 'public, max-age=600';
    // кэш CDN Netlify: одинаковые запросы в течение часа не тратят лимит API
    headers['Netlify-CDN-Cache-Control'] = 'public, s-maxage=3600, stale-while-revalidate=1800, durable';
  } else {
    headers['Cache-Control'] = 'no-store';
  }
  return new Response(JSON.stringify(body), { status, headers });
}

async function tp(token, params) {
  const url = new URL(API);
  for (const [k, v] of Object.entries({ sorting: 'price', currency: 'rub', unique: 'false', ...params })) url.searchParams.set(k, String(v));
  const key = url.toString();
  const hit = mem.get(key);
  if (hit && Date.now() - hit.at < TTL_MS) return hit.data;
  const resp = await fetch(url, { headers: { 'X-Access-Token': token, 'Accept-Encoding': 'gzip, deflate' } });
  if (resp.status === 429) throw Object.assign(new Error('Превышен лимит запросов к API, попробуйте через минуту'), { status: 429 });
  if (!resp.ok) throw Object.assign(new Error(`API ответил ${resp.status}`), { status: 502 });
  const body = await resp.json();
  if (body.success === false) throw Object.assign(new Error(body.error || 'ошибка API'), { status: 502 });
  const data = body.data || [];
  mem.set(key, { at: Date.now(), data });
  if (mem.size > 500) mem.delete(mem.keys().next().value);
  return data;
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

export default async (req) => {
  const q = new URL(req.url).searchParams;
  const origin = (q.get('origin') || 'MOW').toUpperCase();
  const destination = (q.get('destination') || '').toUpperCase();
  const depart = q.get('depart') || '';
  const ret = q.get('return') || '';

  if (!IATA.test(origin) || !IATA.test(destination)) return json({ error: 'Укажите коды IATA откуда и куда (3 латинские буквы)' }, 400, false);
  if (origin === destination) return json({ error: 'Пункты отправления и назначения совпадают' }, 400, false);
  if (!DAY.test(depart) || (ret && !DAY.test(ret))) return json({ error: 'Даты в формате ГГГГ-ММ-ДД' }, 400, false);
  if (ret && ret < depart) return json({ error: 'Дата возврата раньше даты вылета' }, 400, false);
  const today = new Date().toISOString().slice(0, 10);
  if (depart < today) return json({ error: 'Дата вылета уже прошла' }, 400, false);

  const token = process.env.TRAVELPAYOUTS_TOKEN;
  if (!token) return json({ error: 'На сервере не задан TRAVELPAYOUTS_TOKEN' }, 500, false);

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

    return json({
      origin, destination, depart, return: ret || null, one_way: oneWay,
      offers: exact.map(norm).sort((a, b) => a.price - b.price),
      calendar: [...grid.values()].sort((a, b) => (a.depart + a.return).localeCompare(b.depart + b.return)),
      window: WINDOW,
      fetched_at: new Date().toISOString(),
      source: 'Aviasales Data API: кэш поисков за последние дни',
    });
  } catch (e) {
    return json({ error: e.message }, e.status || 502, false);
  }
};

export const config = { path: '/api/flights' };
