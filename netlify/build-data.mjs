// Сборка данных для сайта (запускается Netlify перед публикацией):
//   data/prices.json, data/hotels/*.json  -> site/data/
//   collector/cities.csv + codes.csv      -> site/data/cities.json (IATA, код «Экспресс», адрес Островка)
// Идентификаторы городов c000.. совпадают с collector/collect.py (порядковый номер строки cities.csv).
//
//   node netlify/build-data.mjs              # для деплоя
//   node netlify/build-data.mjs --sample     # только cities.json в site/sample/ (демо-данные)
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const sample = process.argv.includes('--sample');
const OUT = path.join(ROOT, 'site', sample ? 'sample' : 'data');

function parseCsv(text) {
  const rows = [];
  let row = [], cell = '', q = false;
  text = text.replace(/^﻿/, '');
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
      else if (ch === '"') q = false;
      else cell += ch;
    } else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cell); cell = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += ch;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  const [head, ...body] = rows.filter((r) => r.length > 1 || r[0]);
  return body.map((r) => Object.fromEntries(head.map((h, i) => [h, r[i] ?? ''])));
}

const read = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8');
const cities = parseCsv(read('collector/cities.csv'));
const codes = fs.existsSync(path.join(ROOT, 'collector/codes.csv'))
  ? new Map(parseCsv(read('collector/codes.csv')).map((r) => [`${r.city}|${r.country}`, r]))
  : new Map();

const list = cities.map((c, i) => {
  const k = codes.get(`${c.city}|${c.country}`) || {};
  return {
    id: `c${String(i).padStart(3, '0')}`,
    city: c.city,
    country: c.country,
    region: c.region,
    abroad: c.abroad === '1',
    country_code: c.country_code,
    iata: c.iata || null,
    express: k.express || null,
    slug: k.ostrovok_slug || null,
    lat: c.lat ? Number(c.lat) : null,
    lon: c.lon ? Number(c.lon) : null,
  };
});

fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(path.join(OUT, 'cities.json'), JSON.stringify(list));
console.log(`cities.json: ${list.length} городов -> ${path.relative(ROOT, OUT)}`);

if (!sample) {
  const src = path.join(ROOT, 'data');
  let copied = 0;
  if (fs.existsSync(path.join(src, 'prices.json'))) {
    fs.copyFileSync(path.join(src, 'prices.json'), path.join(OUT, 'prices.json'));
    copied++;
  } else {
    console.log('data/prices.json нет — сайт покажет демо-данные из site/sample/');
  }
  const hotelsSrc = path.join(src, 'hotels');
  if (fs.existsSync(hotelsSrc)) {
    fs.mkdirSync(path.join(OUT, 'hotels'), { recursive: true });
    for (const f of fs.readdirSync(hotelsSrc).filter((f) => f.endsWith('.json'))) {
      fs.copyFileSync(path.join(hotelsSrc, f), path.join(OUT, 'hotels', f));
      copied++;
    }
  }
  console.log(`скопировано файлов данных: ${copied}`);
}
