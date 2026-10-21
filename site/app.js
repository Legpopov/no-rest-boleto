'use strict';
// no-rest-boleto: таблица городов, поиск авиа и поездов, карта отелей, итоговая оценка поездки.
// Данные снимка: data/prices.json (если нет — демо из sample/). Живые запросы: /api/flights, /api/trains.

const MOSCOW = { city: 'Москва', iata: 'MOW', express: '2000000' };
const S = {
  base: 'data/', demo: false, prices: null, cities: [], byId: new Map(), rows: [],
  origin: MOSCOW, dest: null, depart: '', ret: '',
  flights: null, flight: null,
  trains: null, train: { there: null, back: null },
  hotelsFor: null, hotels: [], filtered: [], hotel: null,
  nightsTouched: false,
};

// ---------- утилиты ----------
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const rub = (n) => (n == null || Number.isNaN(n) ? '—' : `${Math.round(n).toLocaleString('ru-RU')} ₽`);
const addDays = (iso, n) => { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10); };
const daysBetween = (a, b) => Math.round((Date.parse(b) - Date.parse(a)) / 864e5);
const WD = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
const wd = (iso) => WD[new Date(iso + 'T00:00:00Z').getUTCDay()];
const dm = (iso) => `${iso.slice(8, 10)}.${iso.slice(5, 7)}`;
const dmy = (iso) => `${iso.slice(8, 10)}.${iso.slice(5, 7)}.${iso.slice(0, 4)}`;
const todayIso = () => { const d = new Date(); return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())).toISOString().slice(0, 10); };
const hm = (min) => (min == null ? '' : `${Math.floor(min / 60)} ч ${String(min % 60).padStart(2, '0')} мин`);
const store = {
  get(k) { try { return JSON.parse(localStorage.getItem('nrb:' + k)); } catch { return null; } },
  set(k, v) { try { localStorage.setItem('nrb:' + k, JSON.stringify(v)); } catch { /* приватный режим */ } },
};
function quant(vals, p) {
  if (!vals.length) return null;
  const v = [...vals].sort((a, b) => a - b), k = (v.length - 1) * p, lo = Math.floor(k), hi = Math.min(lo + 1, v.length - 1);
  return v[lo] + (v[hi] - v[lo]) * (k - lo);
}
async function getJson(url) {
  const r = await fetch(url);
  const body = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(body.error || `HTTP ${r.status}`);
  return body;
}

const AIRLINES = {
  SU: 'Аэрофлот', S7: 'S7 Airlines', U6: 'Уральские авиалинии', DP: 'Победа', FV: 'Россия', UT: 'ЮТэйр', N4: 'Северный ветер',
  '5N': 'Смартавиа', A4: 'Азимут', WZ: 'Red Wings', YC: 'Ямал', R3: 'Якутия', IO: 'ИрАэро', '6R': 'АЛРОСА', KV: 'КрасАвиа',
  EO: 'Икар', ZF: 'Azur Air', '7R': 'РусЛайн', Y7: 'NordStar', HZ: 'Аврора', '2S': 'Southwind', B2: 'Белавиа', EK: 'Emirates',
  FZ: 'flydubai', TK: 'Turkish Airlines', PC: 'Pegasus', QR: 'Qatar Airways', EY: 'Etihad', HY: 'Uzbekistan Airways',
  KC: 'Air Astana', J2: 'AZAL', CA: 'Air China', MU: 'China Eastern', CZ: 'China Southern', HU: 'Hainan', AI: 'Air India',
  VN: 'Vietnam Airlines', G9: 'Air Arabia', W5: 'Mahan Air', MS: 'EgyptAir', ET: 'Ethiopian', A9: 'Georgian Airways',
};
const RAIL_CLS = { plazcard: 'плацкарт', coupe: 'купе', sedentary: 'сидячий', lux: 'СВ', soft: 'люкс' };
const RAIL_ORDER = ['sedentary', 'plazcard', 'coupe', 'lux', 'soft'];
// типы Островка приходят как HOTEL, APART_HOTEL, MINI_HOTEL…; ключ — без регистра и разделителей
const KINDS = {
  HOTEL: 'отель', MINIHOTEL: 'мини-отель', HOSTEL: 'хостел', APARTMENT: 'апартаменты', APARTHOTEL: 'апарт-отель',
  GUESTHOUSE: 'гостевой дом', BNB: 'B&B', SANATORIUM: 'санаторий', RESORT: 'курортный отель', BOUTIQUE: 'бутик-отель',
  BOUTIQUEANDDESIGN: 'бутик-отель', COTTAGE: 'коттедж', COTTAGESANDHOUSES: 'коттеджи и дома', CAMPING: 'кемпинг', FARM: 'ферма',
  VILLA: 'вилла', VILLASANDBUNGALOWS: 'виллы и бунгало', GLAMPING: 'глэмпинг', CASTLE: 'замок', CAPSULEHOTEL: 'капсульный отель',
  INN: 'гостиница', LODGE: 'лодж',
};
const kindKey = (k) => String(k || '').toUpperCase().replace(/[^A-Z]/g, '');
const kindLabel = (k) => KINDS[kindKey(k)] || (k ? String(k).replace(/_/g, ' ').toLowerCase() : 'не указан');
const isHostel = (h) => /HOSTEL|CAPSULE/.test(kindKey(h.kind));
const isApart = (h) => h.apart || kindKey(h.kind) === 'APARTMENT';

// ---------- загрузка ----------
async function loadData() {
  let prices;
  try { prices = await getJson('data/prices.json'); } catch { S.base = 'sample/'; S.demo = true; prices = await getJson('sample/prices.json'); }
  let cities;
  try { cities = await getJson(S.base + 'cities.json'); } catch { cities = await getJson('sample/cities.json'); }
  S.prices = prices;
  S.cities = cities;
  for (const c of cities) S.byId.set(c.id, c);
  const pRows = new Map((prices.cities || []).map((r) => [r.id, r]));
  // в таблице — все города списка; данные снимка подставляются, если есть
  S.rows = cities.map((c) => ({ ...c, ...(pRows.get(c.id) || {}), iata: c.iata, express: c.express, slug: c.slug }));
  for (const [id, r] of pRows) if (!S.byId.has(id)) S.rows.push(r);
  if (S.demo) S.rows = S.rows.filter((r) => pRows.has(r.id));
}

function renderSnapshot() {
  const el = $('#snapshot');
  const g = S.prices.generated_at;
  if (!g) { el.textContent = 'дата снимка неизвестна'; return; }
  const d = new Date(g);
  const age = Math.floor((Date.now() - d) / 864e5);
  const when = d.toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'Europe/Moscow' });
  const cov = S.prices.coverage;
  el.innerHTML = `<i class="dot"></i>${S.demo ? '<span class="demo">демо-данные</span> ' : ''}Снимок цен: <b>${when} МСК</b>` +
    ` · ${age <= 0 ? 'сегодня' : age === 1 ? 'вчера' : `${age} дн. назад`}`;
  if (cov) {
    $('#c-stats').innerHTML = [['Городов в снимке', cov.total], ['С ценами на авиа', cov.air], ['С поездами', cov.rail], ['С отелями', cov.hotels]]
      .map(([t, n]) => `<div class="stat"><div class="stat-t">${t}</div><div class="stat-n">${n}</div></div>`).join('');
  }
  el.classList.toggle('old', age > 3);
  el.title = 'Время сбора data/prices.json (generated_at). Живой поиск по датам на вкладках «Авиа» и «Поезда» идёт отдельно.';
  const fx = S.prices.fx;
  if (fx && fx.rates) {
    $('#c-fx').textContent = `Курсы ЦБ РФ на ${fx.date}: ` + ['USD', 'EUR', 'CNY', 'TRY', 'AED', 'KZT'].filter((k) => fx.rates[k]).map((k) => `${k} ${fx.rates[k].toLocaleString('ru-RU', { maximumFractionDigits: 4 })} ₽`).join(' · ');
  }
}

// ---------- параметры поездки ----------
function cityLabel(c) { return c.city; }
function findCity(text) {
  const t = (text || '').trim().toLowerCase().replace(/ё/g, 'е');
  if (!t) return null;
  if (t === 'москва' || t === 'mow') return MOSCOW;
  const name = t.replace(/\s*\(.*\)\s*$/, '');
  const hit = S.cities.find((c) => c.city.toLowerCase().replace(/ё/g, 'е') === name) ||
    S.cities.find((c) => c.iata && c.iata.toLowerCase() === t);
  if (hit) return hit;
  if (/^[a-z]{3}$/.test(t)) return { city: t.toUpperCase(), iata: t.toUpperCase(), express: null, custom: true };
  return null;
}
function fillDatalist() {
  const opts = [`<option value="Москва">Россия · MOW</option>`];
  for (const c of [...S.cities].sort((a, b) => a.city.localeCompare(b.city, 'ru'))) {
    opts.push(`<option value="${esc(cityLabel(c))}">${esc(c.country)}${c.iata ? ' · ' + c.iata : ''}</option>`);
  }
  $('#dl-cities').innerHTML = opts.join('');
}
function readQuery() {
  const q = new URLSearchParams(location.search);
  const saved = store.get('trip') || {};
  const to = q.get('to') || saved.to;
  const from = q.get('from') || saved.from;
  S.dest = (to && (S.byId.get(to) || findCity(to))) || null;
  S.origin = (from && (S.byId.get(from) || findCity(from))) || MOSCOW;
  const t = todayIso();
  S.depart = q.get('d') || saved.d || addDays(t, 14);
  if (S.depart < t) S.depart = addDays(t, 14);
  S.ret = q.has('r') ? q.get('r') : saved.r ?? addDays(S.depart, 3);
  if (S.ret && S.ret < S.depart) S.ret = addDays(S.depart, 3);
}
function writeQuery() {
  const q = new URLSearchParams();
  const key = (c) => (c ? c.id || c.iata : '');
  if (S.dest) q.set('to', key(S.dest));
  if (S.origin !== MOSCOW) q.set('from', key(S.origin));
  q.set('d', S.depart);
  q.set('r', S.ret || '');
  history.replaceState(null, '', `?${q}${location.hash}`);
  store.set('trip', { to: key(S.dest), from: S.origin === MOSCOW ? '' : key(S.origin), d: S.depart, r: S.ret });
}
function syncForm() {
  $('#f-origin').value = S.origin ? cityLabel(S.origin) : '';
  $('#f-dest').value = S.dest ? cityLabel(S.dest) : '';
  $('#f-depart').value = S.depart;
  $('#f-depart').min = todayIso();
  $('#f-return').value = S.ret || '';
  $('#f-return').min = S.depart;
  const n = S.ret ? daysBetween(S.depart, S.ret) : null;
  $('#f-nights').textContent = n == null ? 'в одну сторону' : `${n} ${plural(n, 'ночь', 'ночи', 'ночей')}`;
  if (!S.nightsTouched) $('#t-nights').value = n ?? 0;
}
function plural(n, one, few, many) {
  const m10 = n % 10, m100 = n % 100;
  return m10 === 1 && m100 !== 11 ? one : m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14) ? few : many;
}
function onTripChange(what) {
  const o = findCity($('#f-origin').value);
  const d = findCity($('#f-dest').value);
  const prevDest = S.dest, prevOrigin = S.origin, prevDates = S.depart + S.ret;
  if (o) S.origin = o;
  if (d) S.dest = d;
  S.depart = $('#f-depart').value || S.depart;
  S.ret = $('#f-return').value;
  if (S.ret && S.ret < S.depart) S.ret = S.depart;
  if (what === 'dates') S.nightsTouched = false;
  const routeChanged = prevDest !== S.dest || prevOrigin !== S.origin;
  if (routeChanged || prevDates !== S.depart + S.ret) { S.flights = null; S.flight = null; }
  if (routeChanged) { S.trains = null; S.train = { there: null, back: null }; }
  if (prevDest !== S.dest) S.hotel = null;
  syncForm();
  writeQuery();
  refreshTab();
}

// ---------- вкладки ----------
const TABS = ['cities', 'air', 'rail', 'hotels', 'total'];
let cur = 'cities';
function showTab() {
  cur = TABS.includes(location.hash.slice(1)) ? location.hash.slice(1) : 'cities';
  for (const t of TABS) $(`#tab-${t}`).hidden = t !== cur;
  $$('.tabs a').forEach((a) => a.classList.toggle('active', a.dataset.tab === cur));
  refreshTab();
}
function refreshTab() {
  if (cur === 'cities') renderCities();
  if (cur === 'air') renderAir();
  if (cur === 'rail') renderRail();
  if (cur === 'hotels') renderHotels();
  if (cur === 'total') renderTotal();
}
const needDest = (el) => { el.innerHTML = '<p class="muted">Выберите город в поле «Куда».</p>'; };

// ---------- города ----------
let sortKey = 'city', sortAsc = true;
function railMin(r) {
  const there = r.rail?.status === 'ok' ? r.rail.there || [] : [];
  let best = null;
  for (const t of there) for (const [cls, p] of Object.entries(t[3] || {})) if (best == null || p < best.p) best = { p, cls };
  return best;
}
function cityMetrics(r) {
  const mode = $('#c-airmode').value;
  const a = r.air?.status === 'ok' ? r.air[mode] : null;
  const h = r.hotels?.status === 'ok' ? r.hotels.price_all : null;
  const rm = railMin(r);
  return { city: r.city, country: r.country + (r.region ? ' ' + r.region : ''), air_med: a?.median, air_p90: a?.p90, air_n: a?.n, rail: rm?.p, rail_cls: rm?.cls, hot_med: h?.median, hot_p25: h?.p25, hot_n: h?.n };
}
function renderCities() {
  const q = $('#c-search').value.trim().toLowerCase().replace(/ё/g, 'е');
  const ru = $('#c-abroad-ru').checked, ab = $('#c-abroad-x').checked;
  const list = S.rows
    .filter((r) => (r.abroad ? ab : ru))
    .filter((r) => !q || `${r.city} ${r.country} ${r.region || ''} ${r.iata || ''}`.toLowerCase().replace(/ё/g, 'е').includes(q))
    .map((r) => ({ r, m: cityMetrics(r) }));
  list.sort((x, y) => {
    const a = x.m[sortKey], b = y.m[sortKey];
    if (a == null && b == null) return 0;
    if (a == null) return 1;
    if (b == null) return -1;
    const c = typeof a === 'string' ? a.localeCompare(b, 'ru') : a - b;
    return sortAsc ? c : -c;
  });
  $$('#c-table th').forEach((th) => { th.classList.toggle('sorted', th.dataset.sort === sortKey); th.classList.toggle('asc', th.dataset.sort === sortKey && sortAsc); });
  const stale = (src) => (src?.stale ? ` <span class="sub" title="Последняя удачная загрузка ${esc(src.updated_at)}">(устар.)</span>` : '');
  const status = (src) => (!src ? '' : src.status === 'ok' ? '' : ({ no_airport: 'нет аэропорта', no_station: 'нет ж/д', no_slug: 'нет данных', empty: 'нет цен', blocked: 'не собрано', error: 'ошибка' }[src.status] || src.status));
  $('#c-table tbody').innerHTML = list.map(({ r, m }) => `
    <tr data-id="${esc(r.id)}">
      <td class="c-city"><b>${esc(r.city)}</b>${r.iata ? ` <span class="sub">${esc(r.iata)}</span>` : ''}</td>
      <td class="sub c-country">${esc(m.country)}</td>
      <td class="num" data-label="Авиа, медиана">${m.air_med != null ? rub(m.air_med) + stale(r.air) : `<span class="sub">${status(r.air)}</span>`}</td>
      <td class="num" data-label="Авиа, 90%">${rub(m.air_p90)}</td>
      <td class="num sub c-n">${m.air_n ?? ''}</td>
      <td class="num" data-label="Поезд от">${m.rail != null ? rub(m.rail) + ` <span class="sub">${RAIL_CLS[m.rail_cls] || ''}</span>` + stale(r.rail) : `<span class="sub">${status(r.rail)}</span>`}</td>
      <td class="num" data-label="Отель, медиана">${m.hot_med != null ? rub(m.hot_med) + stale(r.hotels) : `<span class="sub">${status(r.hotels)}</span>`}</td>
      <td class="num" data-label="Отель, 25%">${rub(m.hot_p25)}</td>
      <td class="num sub c-n">${m.hot_n ?? ''}</td>
    </tr>`).join('') || '<tr><td colspan="9" class="muted">Ничего не найдено</td></tr>';
}

// ---------- авиа ----------
const tmode = () => ({ 9: 'any', 1: 'le1', 0: 'direct' }[$('#a-transfers').value]);
async function searchAir() {
  const st = $('#a-status');
  if (!S.dest) { needDest($('#a-results')); return; }
  if (!S.dest.iata) { st.innerHTML = '<span class="err">У этого города нет аэропорта в справочнике.</span>'; return; }
  if (!S.origin.iata) { st.innerHTML = '<span class="err">У города отправления нет аэропорта.</span>'; return; }
  st.textContent = 'ищем…';
  const q = new URLSearchParams({ origin: S.origin.iata, destination: S.dest.iata, depart: S.depart });
  if (S.ret) q.set('return', S.ret);
  try {
    S.flights = await getJson(`/api/flights?${q}`);
    st.textContent = `обновлено ${new Date(S.flights.fetched_at).toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' })}`;
  } catch (e) {
    S.flights = null;
    st.innerHTML = `<span class="err">${esc(e.message)}</span>`;
  }
  renderAir();
}
function renderAir() {
  const res = $('#a-results'), flex = $('#a-flex');
  if (!S.dest) { flex.innerHTML = ''; needDest(res); return; }
  const F = S.flights;
  if (!F) {
    flex.innerHTML = '';
    res.innerHTML = `<p class="muted">${esc(S.origin.city)} → ${esc(S.dest.city)}, ${dmy(S.depart)}${S.ret ? ' — ' + dmy(S.ret) : ''}. Нажмите «Найти билеты».</p>`;
    return;
  }
  const mode = tmode();
  const maxTr = Number($('#a-transfers').value);
  const offers = F.offers.filter((o) => (o.transfers || 0) <= maxTr && (F.one_way || (o.return_transfers || 0) <= maxTr));
  renderFlex(F, mode, offers);
  if (!offers.length) {
    res.innerHTML = `<p class="muted">На точные даты в кэше Aviasales вариантов нет${maxTr < 9 ? ' с такими пересадками' : ''}. Посмотрите соседние даты выше или <a href="${aviasalesSearchUrl()}" target="_blank" rel="noopener">живой поиск на Aviasales</a>.</p>`;
    return;
  }
  res.innerHTML = `<h4>${offers.length} ${plural(offers.length, 'вариант', 'варианта', 'вариантов')} на ${dmy(F.depart)}${F.return ? ' — ' + dmy(F.return) : ''}</h4><div class="offers">` +
    offers.map((o, i) => {
      const leg = (at, tr, dur, from, to) => `${at ? at.slice(11, 16) : ''} ${esc(from || '')}→${esc(to || '')} · ${tr ? `${tr} ${plural(tr, 'пересадка', 'пересадки', 'пересадок')}` : '<span class="tag good">прямой</span>'}${dur ? ' · ' + hm(dur) : ''}`;
      const picked = S.flight && S.flight.link === o.link && S.flight.price === o.price;
      return `<div class="offer${picked ? ' picked' : ''}">
        <img src="https://pics.avs.io/72/72/${encodeURIComponent(o.airline || '')}.png" alt="" loading="lazy" onerror="this.style.visibility='hidden'">
        <div>
          <div><b>${esc(AIRLINES[o.airline] || o.airline || '')}</b> <span class="sub muted">${esc(o.airline || '')} ${esc(o.flight_number || '')}</span></div>
          <div class="legs">Туда ${dm(o.depart_at || F.depart)}: ${leg(o.depart_at, o.transfers, o.duration_to, o.origin_airport, o.destination_airport)}</div>
          ${F.one_way ? '' : `<div class="legs">Обратно ${dm(o.return_at || F.return)}: ${leg(o.return_at, o.return_transfers, o.duration_back, o.destination_airport, o.origin_airport)}</div>`}
        </div>
        <div class="right">
          <div class="price">${rub(o.price)}</div>
          <div class="acts">
            ${o.link ? `<a class="btn buy" href="${esc(o.link)}" target="_blank" rel="noopener">Купить</a>` : ''}
            <button class="btn tint${picked ? ' on' : ''}" data-pick-flight="${i}" type="button">${picked ? 'В итоге ✓' : 'В итог'}</button>
          </div>
        </div>
      </div>`;
    }).join('') + '</div>';
  $$('[data-pick-flight]', res).forEach((b) => b.addEventListener('click', () => {
    S.flight = { ...offers[Number(b.dataset.pickFlight)], route: `${S.origin.city} → ${S.dest.city}` };
    renderAir();
  }));
}
function aviasalesSearchUrl() {
  const d = (iso) => iso.slice(8, 10) + iso.slice(5, 7);
  return `https://www.aviasales.ru/search/${S.origin.iata}${d(S.depart)}${S.dest.iata}${S.ret ? d(S.ret) : ''}1`;
}
function renderFlex(F, mode, offers) {
  const el = $('#a-flex');
  const cal = new Map(F.calendar.map((c) => [`${c.depart}|${c.return}`, c[mode]]));
  const exactMin = offers.length ? Math.min(...offers.map((o) => o.price)) : null;
  const price = (d, r) => {
    const k = `${d}|${F.one_way ? null : r}`;
    const v = cal.get(k);
    if (d === F.depart && (F.one_way || r === F.return)) return exactMin ?? v ?? null;
    return v ?? null;
  };
  const W = F.window;
  const base = price(F.depart, F.return);
  const shifts = [...Array(2 * W + 1).keys()].map((i) => i - W);
  const hints = [];
  const fmtHint = (what, from, to, p) => {
    if (base == null) return `${what} ${dm(to)} (${wd(to)}) — от <b>${rub(p)}</b>`;
    return `Если ${what} <b>${dm(to)}</b> (${wd(to)}) вместо ${dm(from)} — дешевле на <b>${rub(base - p)}</b> (${rub(p)})`;
  };
  // вылет: меняем только дату вылета
  const depAlts = shifts.map((k) => addDays(F.depart, k)).filter((d) => d >= todayIso() && (F.one_way || d < F.return)).map((d) => ({ d, p: price(d, F.return) }));
  const bestDep = depAlts.filter((x) => x.p != null && x.d !== F.depart).sort((a, b) => a.p - b.p)[0];
  if (bestDep && (base == null || bestDep.p < base)) hints.push(fmtHint('вылететь', F.depart, bestDep.d, bestDep.p));
  let retAlts = [], bestRet = null;
  if (!F.one_way) {
    retAlts = shifts.map((k) => addDays(F.return, k)).filter((r) => r > F.depart).map((r) => ({ r, p: price(F.depart, r) }));
    bestRet = retAlts.filter((x) => x.p != null && x.r !== F.return).sort((a, b) => a.p - b.p)[0];
    if (bestRet && (base == null || bestRet.p < base)) hints.push(fmtHint('вернуться', F.return, bestRet.r, bestRet.p));
    // обе даты сразу
    let bestBoth = null;
    for (const a of shifts) for (const b of shifts) {
      const d = addDays(F.depart, a), r = addDays(F.return, b);
      if (d < todayIso() || r <= d) continue;
      const p = price(d, r);
      if (p != null && (!bestBoth || p < bestBoth.p)) bestBoth = { d, r, p };
    }
    if (bestBoth && (bestBoth.d !== F.depart || bestBoth.r !== F.return) && (base == null || bestBoth.p < base) &&
        !(bestDep && bestDep.p <= bestBoth.p) && !(bestRet && bestRet.p <= bestBoth.p)) {
      hints.push(`Самый дешёвый вариант в окне ±${W} дня: <b>${dm(bestBoth.d)}–${dm(bestBoth.r)}</b> — ${rub(bestBoth.p)}${base != null ? `, дешевле на <b>${rub(base - bestBoth.p)}</b>` : ''}`);
    }
  }
  let head;
  if (base == null) head = `<div class="hint">На ${dm(F.depart)}${F.return ? '–' + dm(F.return) : ''} в кэше цен нет.${hints.length ? ' Ближайшие найденные варианты:<br>' + hints.join('<br>') : ' Попробуйте другие даты.'}</div>`;
  else if (hints.length) head = `<div class="hint">${hints.join('<br>')}</div>`;
  else head = `<div class="hint">Выбранные даты — самые дешёвые в окне ±${W} дня (${rub(base)}).</div>`;

  const strip = (label, items, isCur, onPick) => {
    const ps = items.map((x) => x.p).filter((p) => p != null);
    const min = ps.length ? Math.min(...ps) : null;
    return `<div class="strip"><div class="lab">${label}</div>${items.map((x) => {
      const day = x.d || x.r;
      const cls = ['cell', x.p == null ? 'none' : '', isCur(day) ? 'cur' : '', x.p != null && x.p === min ? 'best' : ''].join(' ');
      return `<div class="${cls}" ${x.p != null ? `data-${onPick}="${day}"` : ''} title="${dmy(day)}"><div class="d">${dm(day)} ${wd(day)}</div><div class="p">${x.p != null ? Math.round(x.p / 100) / 10 + 'к' : '—'}</div></div>`;
    }).join('')}</div>`;
  };
  let html = head + strip('Вылет', depAlts, (d) => d === F.depart, 'dep');
  if (!F.one_way) html += strip('Возврат', retAlts, (r) => r === F.return, 'ret');

  if (!F.one_way) {
    const deps = shifts.map((k) => addDays(F.depart, k)).filter((d) => d >= todayIso());
    const rets = shifts.map((k) => addDays(F.return, k));
    const all = [];
    for (const d of deps) for (const r of rets) if (r > d) { const p = price(d, r); if (p != null) all.push(p); }
    if (all.length > 2) {
      const mn = Math.min(...all), mx = Math.max(...all);
      const shade = (p) => `hsla(${Math.round(130 - 130 * ((p - mn) / (mx - mn || 1)))}, 70%, 45%, .22)`;
      html += `<details><summary class="muted">Все сочетания дат ±${W} дня (вылет × возврат)</summary><div class="matrix-wrap"><table class="matrix"><tr><th>вылет ↓ / возврат →</th>${rets.map((r) => `<th>${dm(r)} ${wd(r)}</th>`).join('')}</tr>` +
        deps.map((d) => `<tr><th>${dm(d)} ${wd(d)}</th>${rets.map((r) => {
          const p = r > d ? price(d, r) : null;
          if (p == null) return '<td class="none">—</td>';
          const c = [d === F.depart && r === F.return ? 'cur' : '', p === mn ? 'best' : ''].join(' ');
          return `<td class="${c}" style="background:${shade(p)}" data-both="${d}|${r}">${Math.round(p / 100) / 10}к</td>`;
        }).join('')}</tr>`).join('') + '</table></div></details>';
    }
  }
  el.innerHTML = `<div class="flex">${html}<p class="note">Клик по дню — поиск на эту дату. «к» — тысяч рублей; ${{ any: 'с любыми пересадками', le1: 'не более 1 пересадки', direct: 'только прямые' }[mode]}.</p></div>`;
  const go = (d, r) => { $('#f-depart').value = d; if (r !== undefined) $('#f-return').value = r; onTripChange('dates'); searchAir(); };
  $$('[data-dep]', el).forEach((c) => c.addEventListener('click', () => go(c.dataset.dep)));
  $$('[data-ret]', el).forEach((c) => c.addEventListener('click', () => go(F.depart, c.dataset.ret)));
  $$('[data-both]', el).forEach((c) => c.addEventListener('click', () => { const [d, r] = c.dataset.both.split('|'); go(d, r); }));
}

// ---------- поезда ----------
const tutuUrl = (from, to, iso) => `https://www.tutu.ru/poezda/rasp_d.php?nnst1=${from}&nnst2=${to}${iso ? '&date=' + dmy(iso) : ''}`;
async function searchRail() {
  const st = $('#r-status');
  if (!S.dest) { needDest($('#r-results')); return; }
  const from = S.origin.express, to = S.dest.express;
  if (!from || !to) { st.innerHTML = `<span class="err">Для ${esc(!to ? S.dest.city : S.origin.city)} нет кода ж/д станции в справочнике.</span>`; return; }
  st.textContent = 'загружаем…';
  try {
    const [there, back] = await Promise.all([getJson(`/api/trains?from=${from}&to=${to}`), getJson(`/api/trains?from=${to}&to=${from}`)]);
    S.trains = { there, back, key: `${from}|${to}` };
    st.textContent = '';
  } catch (e) {
    S.trains = null;
    st.innerHTML = `<span class="err">${esc(e.message)}</span>`;
  }
  renderRail();
}
function renderRail() {
  const res = $('#r-results'), links = $('#r-links');
  if (!S.dest) { links.innerHTML = ''; needDest(res); return; }
  const from = S.origin.express, to = S.dest.express;
  links.innerHTML = from && to
    ? `<a href="${tutuUrl(from, to, S.depart)}" target="_blank" rel="noopener">tutu.ru: ${esc(S.origin.city)} → ${esc(S.dest.city)}, ${dmy(S.depart)}</a>` +
      (S.ret ? `<a href="${tutuUrl(to, from, S.ret)}" target="_blank" rel="noopener">tutu.ru: обратно, ${dmy(S.ret)}</a>` : '')
    : '';
  if (!S.trains || S.trains.key !== `${from}|${to}`) {
    res.innerHTML = from && to ? '<p class="muted">Нажмите «Показать поезда».</p>' : `<p class="muted">Для этого направления нет кода ж/д станции.</p>`;
    // данные снимка: подскажем, есть ли поезда
    return;
  }
  const firmOnly = $('#r-firm').checked;
  const table = (dir, data) => {
    const trips = data.trips.filter((t) => !firmOnly || t.firm);
    if (!trips.length) return `<p class="muted">Поездов не найдено${firmOnly ? ' (с фильтром «фирменные»)' : ''}.</p>`;
    let cheapest = null;
    for (const t of trips) for (const [c, p] of Object.entries(t.classes)) if (!cheapest || p < cheapest.p) cheapest = { p, c, t: t.train };
    return `<p class="muted">Самый дешёвый: ${rub(cheapest.p)} — ${RAIL_CLS[cheapest.c] || cheapest.c}, поезд ${esc(cheapest.t)}</p>
    <div class="table-scroll"><table class="grid cards"><thead><tr><th>Поезд</th><th>Отпр.</th><th>Приб.</th><th class="num">В пути</th><th>Цены от (нажмите, чтобы взять в итог)</th></tr></thead><tbody>` +
      trips.map((t) => `<tr>
        <td><b>${esc(t.train)}</b>${t.name ? ` <span class="sub">${esc(t.name)}</span>` : ''}${t.firm ? ' <span class="tag">фирменный</span>' : ''}</td>
        <td data-label="Отправление">${esc(t.departure)}</td><td data-label="Прибытие">${esc(t.arrival)}</td><td class="num" data-label="В пути">${t.hours} ч</td>
        <td class="t-cls"><div class="train-cls">${RAIL_ORDER.filter((c) => t.classes[c]).map((c) => {
          const on = S.train[dir] && S.train[dir].train === t.train && S.train[dir].cls === c;
          return `<button type="button" class="${on ? 'on' : ''}" data-dir="${dir}" data-train="${esc(t.train)}" data-cls="${c}" data-price="${t.classes[c]}">${RAIL_CLS[c]} ${rub(t.classes[c])}</button>`;
        }).join('')}</div></td></tr>`).join('') + '</tbody></table></div>';
  };
  res.innerHTML = `<h4>Туда: ${esc(S.origin.city)} → ${esc(S.dest.city)}</h4>${table('there', S.trains.there)}` +
    `<h4>Обратно: ${esc(S.dest.city)} → ${esc(S.origin.city)}</h4>${table('back', S.trains.back)}` +
    `<p class="note">Источник: ${esc(S.trains.there.source)}. Время отправления и прибытия — местное по расписанию.</p>`;
  $$('[data-train]', res).forEach((b) => b.addEventListener('click', () => {
    const { dir, train, cls, price } = b.dataset;
    const curSel = S.train[dir];
    S.train[dir] = curSel && curSel.train === train && curSel.cls === cls ? null : { train, cls, price: Number(price) };
    renderRail();
  }));
}

// ---------- отели ----------
let map = null, layer = null, markers = new Map();
function hotelDates() { return { inn: S.depart, out: S.ret && S.ret > S.depart ? S.ret : addDays(S.depart, 1) }; }
function hotelLinks(c) {
  const { inn, out } = hotelDates();
  const people = Number($('#t-people').value) || 1;
  const L = [];
  if (c.slug) L.push(`<a href="https://ostrovok.ru/hotel/${c.slug}/?dates=${dmy(inn)}-${dmy(out)}&guests=${people}" target="_blank" rel="noopener">Островок на ${dm(inn)}–${dm(out)}</a>`);
  // у Яндекса поиск требует внутренний geoId города, поэтому открываем раздел отелей с датами, город вводится на сайте
  L.push(`<a href="https://travel.yandex.ru/hotels/?adults=${people}&checkinDate=${inn}&checkoutDate=${out}&text=${encodeURIComponent(c.city)}" target="_blank" rel="noopener" title="Город может понадобиться ввести на сайте">Яндекс Путешествия</a>`);
  L.push(`<a href="https://www.booking.com/searchresults.ru.html?ss=${encodeURIComponent(c.city + (c.country ? ', ' + c.country : ''))}&checkin=${inn}&checkout=${out}&group_adults=${people}&no_rooms=1" target="_blank" rel="noopener">Booking</a>`);
  return L.join('');
}
const hotelUrl = (slug, h) => {
  const { inn, out } = hotelDates();
  return h.slug_id && h.id ? `https://ostrovok.ru/hotel/${slug}/mid${h.id}/${encodeURIComponent(h.slug_id)}/?dates=${dmy(inn)}-${dmy(out)}&guests=${Number($('#t-people').value) || 1}` : null;
};
async function loadHotels() {
  const c = S.dest;
  if (S.hotelsFor === c.id) return;
  S.hotelsFor = c.id;
  S.hotels = [];
  S.hotelsMeta = null;
  $('#h-note').textContent = 'загружаем отели…';
  try {
    const j = await getJson(`${S.base}hotels/${c.id}.json`);
    S.hotels = j.hotels || [];
    S.hotelsMeta = j;
  } catch {
    S.hotelsMeta = null;
  }
  buildKindFilters();
}
function buildKindFilters() {
  const kinds = [...new Set(S.hotels.map((h) => h.kind || ''))].sort();
  $('#h-kinds').innerHTML = kinds.map((k) => `<label class="chk"><input type="checkbox" data-kind="${esc(k)}" checked> ${esc(kindLabel(k))} <span class="sub muted">${S.hotels.filter((h) => (h.kind || '') === k).length}</span></label>`).join('') || '<span class="muted">—</span>';
  $$('#h-kinds input').forEach((i) => i.addEventListener('change', applyHotelFilters));
  const maxKm = Math.max(1, ...S.hotels.map((h) => (h.dist_m || 0) / 1000));
  const r = $('#h-dist');
  r.max = Math.min(50, Math.ceil(maxKm * 2) / 2);
  r.value = r.max;
}
function readFilters() {
  return {
    stars: Number($('#h-stars').value), exact: $('#h-stars-exact').checked, rating: Number($('#h-rating').value),
    dist: Number($('#h-dist').value), distMax: Number($('#h-dist').max), pmin: Number($('#h-pmin').value) || 0, pmax: Number($('#h-pmax').value) || Infinity,
    noApart: $('#h-noapart').checked, noHostel: $('#h-nohostel').checked,
    kinds: new Set($$('#h-kinds input:checked').map((i) => i.dataset.kind)), sort: $('#h-sort').value,
  };
}
function applyHotelFilters() {
  const f = readFilters();
  $('#h-dist-out').textContent = f.dist >= f.distMax ? 'без ограничения' : `${f.dist} км`;
  S.filtered = S.hotels.filter((h) =>
    (f.exact ? (h.stars || 0) === f.stars : (h.stars || 0) >= f.stars) &&
    (!f.rating || (h.rating || 0) >= f.rating) &&
    (f.dist >= f.distMax || (h.dist_m != null && h.dist_m <= f.dist * 1000)) &&
    h.price >= f.pmin && h.price <= f.pmax &&
    !(f.noApart && isApart(h)) && !(f.noHostel && isHostel(h)) &&
    f.kinds.has(h.kind || ''));
  const by = { price: (a, b) => a.price - b.price, dist: (a, b) => (a.dist_m ?? 1e9) - (b.dist_m ?? 1e9), rating: (a, b) => (b.rating || 0) - (a.rating || 0) }[f.sort];
  S.filtered.sort(by);
  store.set('hfilters', { stars: f.stars, exact: f.exact, rating: f.rating, noApart: f.noApart, noHostel: f.noHostel, sort: f.sort });
  drawHotels();
}
function drawHotels() {
  const c = S.dest, list = S.filtered, slug = S.hotelsMeta?.slug || c.slug;
  const prices = list.map((h) => h.price);
  const med = quant(prices, 0.5);
  $('#h-stats').innerHTML = list.length
    ? `Подходит: <b>${list.length}</b> из ${S.hotels.length}<br>мин ${rub(Math.min(...prices))} · 25% ${rub(quant(prices, 0.25))}<br>медиана <b>${rub(med)}</b> · 90% ${rub(quant(prices, 0.9))}<br><span class="muted">за ночь, цена «от»</span>`
    : `Подходит: 0 из ${S.hotels.length}`;
  S.hotelMedian = med;
  // карта
  if (map) {
    layer.clearLayers();
    markers.clear();
    const lo = quant(prices, 0.1) ?? 0, hi = quant(prices, 0.9) ?? 1;
    for (const h of list) {
      if (h.lat == null || h.lon == null) continue;
      const t = Math.max(0, Math.min(1, (h.price - lo) / (hi - lo || 1)));
      const picked = S.hotel && S.hotel.id === h.id;
      const m = L.circleMarker([h.lat, h.lon], { radius: picked ? 10 : 7, weight: picked ? 3 : 1, color: picked ? '#2563eb' : '#333', fillColor: `hsl(${Math.round(130 - 130 * t)},70%,45%)`, fillOpacity: 0.85 });
      const url = hotelUrl(slug, h);
      m.bindPopup(`<b>${esc(h.name)}</b><br>${'★'.repeat(h.stars || 0) || 'без звёзд'} · ${esc(kindLabel(h.kind))}${h.rating ? ` · ${h.rating}/10 (${h.reviews})` : ''}<br>${h.dist_m != null ? (h.dist_m / 1000).toFixed(1) + ' км до центра · ' : ''}<b>от ${rub(h.price)}</b>/ночь<br>${url ? `<a href="${esc(url)}" target="_blank" rel="noopener">Открыть на Островке</a> · ` : ''}<a href="#" data-pick-hotel="${esc(h.id)}">в итог</a>`);
      m.on('click', () => highlightHotel(h.id));
      m.addTo(layer);
      markers.set(String(h.id), m);
    }
    const pts = list.filter((h) => h.lat != null).map((h) => [h.lat, h.lon]);
    if (pts.length) map.fitBounds(pts, { padding: [20, 20], maxZoom: 15 });
    else if (c.lat != null) map.setView([c.lat, c.lon], 12);
  }
  // список
  const shown = list.slice(0, 120);
  $('#h-list').innerHTML = shown.map((h) => {
    const url = hotelUrl(slug, h);
    const picked = S.hotel && S.hotel.id === h.id;
    return `<div class="hotel${picked ? ' picked' : ''}" data-hid="${esc(h.id)}">
      <div><b>${esc(h.name)}</b> ${h.stars ? '★'.repeat(h.stars) : ''}
        <div class="meta">${esc(kindLabel(h.kind))}${isApart(h) && kindKey(h.kind) !== 'APARTMENT' ? ' · апартаменты' : ''}${h.rating ? ` · рейтинг ${h.rating} (${h.reviews} отз.)` : ''}${h.dist_m != null ? ` · ${(h.dist_m / 1000).toFixed(1)} км до центра` : ''}${h.address ? ' · ' + esc(h.address) : ''}</div>
      </div>
      <div><div class="price">от ${rub(h.price)}</div>
        <div class="acts" style="display:flex;gap:6px;justify-content:flex-end;flex-wrap:wrap">
          ${url ? `<a class="btn gray" href="${esc(url)}" target="_blank" rel="noopener">Островок</a>` : ''}
          <button class="btn tint${picked ? ' on' : ''}" type="button" data-pick-hotel="${esc(h.id)}">${picked ? 'В итоге ✓' : 'В итог'}</button>
        </div></div>
    </div>`;
  }).join('') + (list.length > shown.length ? `<p class="muted">Показаны первые ${shown.length} из ${list.length}.</p>` : '');
}
function highlightHotel(id) {
  $$('.hotel.hl').forEach((e) => e.classList.remove('hl'));
  const el = $(`.hotel[data-hid="${CSS.escape(String(id))}"]`);
  if (el) { el.classList.add('hl'); el.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); }
}
function pickHotel(id) {
  const h = S.hotels.find((x) => String(x.id) === String(id));
  if (!h) return;
  S.hotel = S.hotel && S.hotel.id === h.id ? null : { ...h, city: S.dest.city };
  drawHotels();
}
async function renderHotels() {
  const links = $('#h-links');
  if (!S.dest || !S.dest.id) {
    links.innerHTML = '';
    $('#h-list').innerHTML = '<p class="muted">Выберите город из списка в поле «Куда».</p>';
    $('#h-note').textContent = '';
    return;
  }
  const c = S.dest;
  links.innerHTML = hotelLinks(c);
  if (!map && window.L) {
    map = L.map('map', { scrollWheelZoom: false });
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>' }).addTo(map);
    layer = L.layerGroup().addTo(map);
    $('#map').addEventListener('click', (e) => { const a = e.target.closest('[data-pick-hotel]'); if (a) { e.preventDefault(); pickHotel(a.dataset.pickHotel); } });
  }
  if (map) { map.invalidateSize(); if (c.lat != null) map.setView([c.lat, c.lon], 12); }
  await loadHotels();
  if (S.dest !== c) return;
  const meta = S.hotelsMeta;
  $('#h-note').innerHTML = meta
    ? `${S.hotels.length} ${plural(S.hotels.length, 'объект', 'объекта', 'объектов')} с публичной выдачи Островка без дат — цена «от» за ночь, обновлено ${new Date(meta.updated_at).toLocaleDateString('ru-RU')}${meta.demo ? ' · <b>демо-данные, не реальные отели</b>' : ''}. На конкретные даты цены другие — откройте ссылку выше.`
    : 'Для этого города в снимке нет списка отелей. Воспользуйтесь ссылками выше — даты уже подставлены.';
  applyHotelFilters();
}

// ---------- итог ----------
function renderTotal() {
  const people = Math.max(1, Number($('#t-people').value) || 1);
  const f = S.flight;
  $('#t-flight').innerHTML = f ? `${esc(f.route)}: ${esc(AIRLINES[f.airline] || f.airline || '')}, ${dm((f.depart_at || '').slice(0, 10) || S.depart)}${f.return_at ? '–' + dm(f.return_at.slice(0, 10)) : ''} — <b>${rub(f.price)}</b> на человека` : 'не выбран — выберите на вкладке «Авиа»';
  const tr = S.train;
  const trPrice = (tr.there?.price || 0) + (tr.back?.price || 0);
  $('#t-train').innerHTML = tr.there || tr.back
    ? [tr.there && `туда ${esc(tr.there.train)}, ${RAIL_CLS[tr.there.cls]} ${rub(tr.there.price)}`, tr.back && `обратно ${esc(tr.back.train)}, ${RAIL_CLS[tr.back.cls]} ${rub(tr.back.price)}`].filter(Boolean).join('; ') + ` — <b>${rub(trPrice)}</b>${!tr.there || !tr.back ? ' (выбрано только одно направление)' : ''}`
    : 'не выбран — выберите на вкладке «Поезда»';
  const h = S.hotel;
  $('#t-hotel-picked').innerHTML = h ? `${esc(h.name)}, ${esc(h.city)} — <b>от ${rub(h.price)}</b>/ночь` : 'не выбран — выберите на вкладке «Отели»';
  $('#t-hotel-median').innerHTML = S.hotelMedian != null && S.hotelsFor === S.dest?.id ? `${esc(S.dest.city)}: ${rub(S.hotelMedian)}/ночь (${S.filtered.length} отелей по фильтрам)` : 'откройте вкладку «Отели» для города';

  const tMode = $('input[name=t-ticket]:checked').value;
  const ticket = tMode === 'flight' ? f?.price : tMode === 'train' ? (trPrice || null) : Number($('#t-ticket-manual').value) || null;
  const hMode = $('input[name=t-hotel]:checked').value;
  const night = hMode === 'picked' ? h?.price : hMode === 'median' ? (S.hotelsFor === S.dest?.id ? S.hotelMedian : null) : Number($('#t-hotel-manual').value) || null;
  const nights = Math.max(0, Number($('#t-nights').value) || 0);
  const extra = Number($('#t-extra').value) || 0;
  const roadSum = ticket != null ? ticket * people : null;
  const hotelSum = night != null ? night * nights : null;
  const total = (roadSum || 0) + (hotelSum || 0) + extra;
  const missing = [roadSum == null && 'дорога', hotelSum == null && nights > 0 && 'проживание'].filter(Boolean);
  $('#t-result').innerHTML = `<h3>${S.dest ? `${esc(S.origin.city)} → ${esc(S.dest.city)}, ` : ''}${dmy(S.depart)}${S.ret ? ' — ' + dmy(S.ret) : ''}</h3>
    <table>
      <tr><td>Дорога${people > 1 ? ` × ${people} чел.` : ''}</td><td class="num">${rub(roadSum)}</td></tr>
      <tr><td>Проживание: ${night != null ? rub(night) : '—'} × ${nights} ${plural(nights, 'ночь', 'ночи', 'ночей')}</td><td class="num">${rub(hotelSum)}</td></tr>
      ${extra ? `<tr><td>Дополнительно</td><td class="num">${rub(extra)}</td></tr>` : ''}
      <tr class="sum"><td>Итого, оценка</td><td class="num">${rub(total)}</td></tr>
    </table>
    <p class="note">${missing.length ? `Не задано: ${missing.join(', ')}. ` : ''}Оценка по рыночным ценам «от»: авиа — кэш поисков Aviasales, поезда и отели — без привязки к датам. Перед покупкой проверьте цены по ссылкам.</p>`;
}

// ---------- запуск ----------
function restoreFilters() {
  const f = store.get('hfilters');
  if (!f) return;
  $('#h-stars').value = String(f.stars ?? 0);
  $('#h-stars-exact').checked = !!f.exact;
  $('#h-rating').value = String(f.rating ?? 0);
  $('#h-noapart').checked = !!f.noApart;
  $('#h-nohostel').checked = !!f.noHostel;
  $('#h-sort').value = f.sort || 'price';
}
async function main() {
  try {
    await loadData();
  } catch (e) {
    $('#snapshot').innerHTML = `<span class="err">Не удалось загрузить данные: ${esc(e.message)}</span>`;
    return;
  }
  renderSnapshot();
  fillDatalist();
  readQuery();
  restoreFilters();
  syncForm();

  $('#f-origin').addEventListener('change', () => onTripChange('route'));
  $('#f-dest').addEventListener('change', () => onTripChange('route'));
  $('#f-depart').addEventListener('change', () => {
    const n = S.ret ? daysBetween(S.depart, S.ret) : null;
    const d = $('#f-depart').value;
    if (n != null && d) $('#f-return').value = addDays(d, n); // сохраняем длительность поездки
    onTripChange('dates');
  });
  $('#f-return').addEventListener('change', () => onTripChange('dates'));
  $('#trip-form').addEventListener('submit', (e) => {
    e.preventDefault();
    onTripChange('route');
    if (!S.dest) return;
    if (cur === 'cities') location.hash = S.dest.iata ? 'air' : S.dest.express ? 'rail' : 'hotels';
    if (location.hash === '#air') searchAir();
    else if (location.hash === '#rail') searchRail();
  });
  // переключатели-«таблетки»: кнопки .seg[data-for] управляют скрытым <select>
  for (const seg of $$('.seg[data-for]')) {
    const sel = $('#' + seg.dataset.for);
    const sync = () => $$('button', seg).forEach((b) => b.classList.toggle('on', b.dataset.value === sel.value));
    seg.addEventListener('click', (e) => {
      const b = e.target.closest('button[data-value]');
      if (!b) return;
      sel.value = b.dataset.value;
      sel.dispatchEvent(new Event('change'));
      sync();
    });
    sync();
  }

  $('#c-search').addEventListener('input', renderCities);
  ['#c-airmode', '#c-abroad-ru', '#c-abroad-x'].forEach((s) => $(s).addEventListener('change', renderCities));
  $$('#c-table th').forEach((th) => th.addEventListener('click', () => {
    if (sortKey === th.dataset.sort) sortAsc = !sortAsc;
    else { sortKey = th.dataset.sort; sortAsc = !['air_n', 'hot_n'].includes(sortKey); }
    renderCities();
  }));
  $('#c-table tbody').addEventListener('click', (e) => {
    const tr = e.target.closest('tr[data-id]');
    if (!tr) return;
    const c = S.byId.get(tr.dataset.id);
    if (!c) return;
    $('#f-dest').value = c.city;
    onTripChange('route');
    location.hash = c.iata ? 'air' : c.express ? 'rail' : 'hotels';
  });

  $('#a-search').addEventListener('click', searchAir);
  $('#a-transfers').addEventListener('change', renderAir);
  $('#r-search').addEventListener('click', searchRail);
  $('#r-firm').addEventListener('change', renderRail);

  ['#h-stars', '#h-stars-exact', '#h-rating', '#h-noapart', '#h-nohostel', '#h-sort'].forEach((s) => $(s).addEventListener('change', applyHotelFilters));
  ['#h-dist', '#h-pmin', '#h-pmax'].forEach((s) => $(s).addEventListener('input', applyHotelFilters));
  $('#h-list').addEventListener('click', (e) => {
    const b = e.target.closest('[data-pick-hotel]');
    if (b) { pickHotel(b.dataset.pickHotel); return; }
    const row = e.target.closest('.hotel');
    if (row && !e.target.closest('a') && markers.get(row.dataset.hid)) {
      const m = markers.get(row.dataset.hid);
      map.panTo(m.getLatLng());
      m.openPopup();
    }
  });

  $('#h-ftoggle').addEventListener('click', () => {
    const open = $('#h-filters').classList.toggle('open');
    $('#h-ftoggle').setAttribute('aria-expanded', String(open));
  });
  $('#t-nights').addEventListener('input', () => { S.nightsTouched = true; renderTotal(); });
  $$('#tab-total input').forEach((i) => i.addEventListener('input', renderTotal));
  $('#t-people').addEventListener('input', () => { if (S.dest) $('#h-links').innerHTML = hotelLinks(S.dest); });

  window.addEventListener('hashchange', showTab);
  showTab();
}
main();
