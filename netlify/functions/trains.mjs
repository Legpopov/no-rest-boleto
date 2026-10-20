// Поезда между станциями «Экспресс»: расписание и минимальные цены по классам вагонов.
// GET /api/trains?from=2000000&to=2060615
// Источник: сервис Туту через suggest.travelpayouts.com. Цены «от», без привязки к дате.

const API = 'https://suggest.travelpayouts.com/search';
const CODE = /^\d{7}$/;
const TTL_MS = 6 * 60 * 60 * 1000;
const mem = new Map();

function json(body, status = 200, cache = true) {
  const headers = { 'Content-Type': 'application/json; charset=utf-8' };
  if (cache) {
    headers['Cache-Control'] = 'public, max-age=3600';
    headers['Netlify-CDN-Cache-Control'] = 'public, s-maxage=21600, stale-while-revalidate=3600, durable';
  } else {
    headers['Cache-Control'] = 'no-store';
  }
  return new Response(JSON.stringify(body), { status, headers });
}

export default async (req) => {
  const q = new URL(req.url).searchParams;
  const from = q.get('from') || '';
  const to = q.get('to') || '';
  if (!CODE.test(from) || !CODE.test(to)) return json({ error: 'Коды станций «Экспресс» — 7 цифр' }, 400, false);

  const url = `${API}?service=tutu_trains&term=${from}&term2=${to}`;
  try {
    let body = mem.get(url);
    if (!body || Date.now() - body.at > TTL_MS) {
      const resp = await fetch(url);
      if (!resp.ok) return json({ error: `Сервис расписаний ответил ${resp.status}` }, 502, false);
      body = { at: Date.now(), data: await resp.json() };
      mem.set(url, body);
      if (mem.size > 500) mem.delete(mem.keys().next().value);
    }
    const trips = (body.data.trips || [])
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
    return json({ from, to, trips, fetched_at: new Date().toISOString(), source: 'Туту (через Travelpayouts), цены от, без привязки к дате' });
  } catch (e) {
    return json({ error: `Сервис расписаний недоступен: ${e.message}` }, 502, false);
  }
};

export const config = { path: '/api/trains' };
