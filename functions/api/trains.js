// Cloudflare Pages Function: /api/trains. Логика — в netlify/lib/api.mjs (общая с Netlify).
import { trains, edgeCached } from '../../netlify/lib/api.mjs';

export const onRequest = (ctx) => edgeCached(ctx, () => trains(new URL(ctx.request.url)), ['from', 'to']);
