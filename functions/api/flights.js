// Cloudflare Pages Function: /api/flights. Логика — в netlify/lib/api.mjs (общая с Netlify).
// Токен: переменная окружения TRAVELPAYOUTS_TOKEN в настройках проекта Pages (локально — файл .dev.vars).
import { flights, edgeCached } from '../../netlify/lib/api.mjs';

export const onRequest = (ctx) => edgeCached(ctx, () => flights(new URL(ctx.request.url), ctx.env.TRAVELPAYOUTS_TOKEN),
  { origin: 'MOW', destination: '', depart: '', return: '' });
