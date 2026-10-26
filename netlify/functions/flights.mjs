// Netlify Function: /api/flights. Логика — в netlify/lib/api.mjs (общая с Cloudflare Pages).
import { flights } from '../lib/api.mjs';

export default (req) => flights(new URL(req.url), process.env.TRAVELPAYOUTS_TOKEN);

export const config = { path: '/api/flights' };
