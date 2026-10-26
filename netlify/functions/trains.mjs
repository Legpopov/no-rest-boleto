// Netlify Function: /api/trains. Логика — в netlify/lib/api.mjs (общая с Cloudflare Pages).
import { trains } from '../lib/api.mjs';

export default (req) => trains(new URL(req.url));

export const config = { path: '/api/trains' };
