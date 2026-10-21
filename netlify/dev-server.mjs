// Локальная проверка без netlify-cli: статика из site/ + функции из netlify/functions по их config.path.
//   node netlify/dev-server.mjs [порт]      (TRAVELPAYOUTS_TOKEN берётся из окружения или из .env)
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SITE = path.join(ROOT, 'site');
const PORT = Number(process.argv[2]) || 8888;

const envFile = path.join(ROOT, '.env');
if (fs.existsSync(envFile)) {
  for (const line of fs.readFileSync(envFile, 'utf8').replace(/^﻿/, '').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^['"]|['"]$/g, '');
  }
}

const routes = {};
for (const f of fs.readdirSync(path.join(ROOT, 'netlify/functions')).filter((f) => f.endsWith('.mjs'))) {
  const mod = await import(pathToFileURL(path.join(ROOT, 'netlify/functions', f)));
  routes[mod.config?.path || `/.netlify/functions/${f.replace(/\.mjs$/, '')}`] = mod.default;
}

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon' };

http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://localhost:${PORT}`);
  const fn = routes[url.pathname];
  try {
    if (fn) {
      const r = await fn(new Request(url, { method: req.method }), {});
      res.writeHead(r.status, Object.fromEntries(r.headers));
      res.end(Buffer.from(await r.arrayBuffer()));
      console.log(r.status, url.pathname + url.search);
      return;
    }
    let file = path.join(SITE, decodeURIComponent(url.pathname));
    if (!file.startsWith(SITE)) throw new Error('bad path');
    if (fs.existsSync(file) && fs.statSync(file).isDirectory()) file = path.join(file, 'index.html');
    if (!fs.existsSync(file)) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
    fs.createReadStream(file).pipe(res);
  } catch (e) {
    res.writeHead(500); res.end(String(e));
  }
}).listen(PORT, () => console.log(`http://localhost:${PORT}  функции: ${Object.keys(routes).join(', ')}`));
