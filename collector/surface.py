"""Independent one-way rail/bus snapshots for explicit branch-origin route scopes.

These are undated/timetable reference prices, never historical monthly means.
Raw HTTP receipts go to --raw-dir, outside the public repository by default.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
from html.parser import HTMLParser
import json
from pathlib import Path
import re
from urllib.parse import urlencode, quote

from .common import PoliteSession
from .planning import positive
from .sources import tutu
from .sources.atlas import NEXT_DATA
from .planning_scope import CITIES
from .branch_scope import endpoints, select_routes


def now():
    return datetime.now(timezone.utc).isoformat()


class ReceiptSession(PoliteSession):
    def __init__(self, raw_dir):
        super().__init__(1.0, 1.8)
        # Some timetable responses declare gzip twice but encode it once.
        # Ask for the ordinary uncompressed HTTP representation instead of
        # guessing how to decode a malformed payload.
        self.headers['Accept-Encoding'] = 'identity'
        # Some timetable responses declare gzip twice but encode it once.
        # Ask for the ordinary uncompressed HTTP representation instead of
        # guessing how to decode a malformed payload.
        self.headers['Accept-Encoding'] = 'identity'
        self.raw_dir = Path(raw_dir)
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self.receipts = []

    def request(self, method, url, *args, **kwargs):
        response = super().request(method, url, *args, **kwargs)
        digest = hashlib.sha256(response.content).hexdigest()
        (self.raw_dir/(digest+'.raw')).write_bytes(response.content)
        self.receipts.append({
            'url': response.url, 'retrieved_at': now(), 'status': response.status_code,
            'sha256': digest, 'bytes': len(response.content),
            'server_date': response.headers.get('Date'),
            'last_modified': response.headers.get('Last-Modified'),
        })
        return response


class BusPage(HTMLParser):
    def __init__(self):
        super().__init__()
        self.scripts, self.text, self.current = [], [], None

    def handle_starttag(self, tag, attrs):
        if tag == 'script' and dict(attrs).get('type') == 'application/ld+json':
            self.current = []

    def handle_data(self, value):
        if self.current is not None:
            self.current.append(value)
        else:
            self.text.append(value)

    def handle_endtag(self, tag):
        if tag == 'script' and self.current is not None:
            self.scripts.append(json.loads(''.join(self.current)))
            self.current = None


def city_key(value):
    return re.sub(r'\s+', ' ', value.casefold().replace('ё','е').replace('_',' ')).strip()


def bus_quotes(html, origin, destination):
    page = BusPage()
    page.feed(html)
    rows, seen = [], set()
    expected = city_key(origin+' – '+destination)
    for root in page.scripts:
        roots = root if isinstance(root, list) else [root]
        for offer in roots:
            if not isinstance(offer, dict) or offer.get('@type') != 'AggregateOffer':
                continue
            if offer.get('priceCurrency') != 'RUB':
                raise ValueError('Bus fares must explicitly use RUB')
            for item in offer.get('offers') or []:
                trip = item.get('itemOffered') or {}
                if city_key(trip.get('name','').replace(' → ',' – ').replace(' — ',' – ')) != expected:
                    raise ValueError('Bus quote belongs to another direction')
                price = item.get('price')
                if not positive(price):
                    continue
                row = {
                    'price': price, 'provider': (trip.get('provider') or {}).get('name'),
                    'number': trip.get('busNumber'), 'departure': trip.get('departureTime'),
                    'departure_stop': (trip.get('departureBusStop') or {}).get('name'),
                    'arrival_stop': (trip.get('arrivalBusStop') or {}).get('name'),
                }
                key = json.dumps(row, ensure_ascii=False, sort_keys=True)
                if key not in seen:
                    rows.append(row)
                    seen.add(key)
    text = re.sub(r'\s+', ' ', ' '.join(page.text)).strip()
    no_direct = 'Прямых автобусов по этому маршруту нет' in text
    if no_direct and rows:
        raise ValueError('Conflicting direct-route status and bus prices')
    return rows, no_direct


def atlas_quotes(html, origin, destination):
    match=NEXT_DATA.search(html)
    if not match:return []
    props=json.loads(match[1]).get('props') or {}
    seo=(props.get('pageProps') or {}).get('seoDirection') or {}
    if not seo:return []
    if (city_key(seo.get('fromName',''))!=city_key(origin)
            or city_key(seo.get('toName',''))!=city_key(destination)):
        raise ValueError('Atlas quote belongs to another direction')
    items=[]
    calendar=((props.get('initialState') or {}).get('searchRides') or {}).get('calendar') or []
    for day in calendar:
        if not day.get('rideCount'):continue
        for price in day.get('minPrices') or []:
            if price.get('currency')!='RUB':raise ValueError('Atlas fare must explicitly use RUB')
            if positive(price.get('price')):
                items.append(dict(price=price['price'],departure=day['date'],provider='Atlas',
                                  statistic='minimum_cached_fare_for_departure_date'))
    # Do not mix undated low/high prices with dated daily minima in the mean.
    if not items:
        for price in seo.get('minPrices') or []:
            if price.get('currency')!='RUB':raise ValueError('Atlas fare must explicitly use RUB')
            if positive(price.get('price')):
                items.append(dict(price=price['price'],provider='Atlas',
                                  statistic='undated_minimum_reference'))
    return items


def atlas_fallback(session, row):
    if row['mode']!='bus' or row['status']=='ok':return row
    if row.get('ambiguous_atlas_endpoint'):
        return row
    previous={k:row.get(k) for k in ('source','status','error')}
    url='https://atlasbus.ru/'+quote('Маршруты')+'/'+quote(row['origin'])+'/'+quote(row['destination'])
    try:
        response=session.get(url);response.raise_for_status()
        items=atlas_quotes(response.text,row['origin'],row['destination'])
        row.setdefault('source_attempts',[]).append(previous)
        if items:
            row.update(source=url,status='ok',items=items,statistic='mean_of_reference_daily_minima_or_undated_minimum')
            row.pop('error',None)
        else:row['source_attempts'].append(dict(source=url,status='no_priced_services'))
    except Exception as error:
        row.setdefault('source_attempts',[]).append(dict(source=url,status='source_error',error_type=type(error).__name__))
    return row


def collect_route(session, origin, destination, mode, registry=None):
    start = len(session.receipts)
    row = {'origin': origin, 'destination': destination, 'mode': mode,
           'currency': 'RUB', 'basis': 'one_way', 'checked_at': now(),
           'price_period': 'undated_reference_not_monthly_history'}
    try:
        registry = registry or {n: {'express':v[0], 'bus_slug':v[1]} for n,v in CITIES.items()}
        if mode == 'rail':
            if not registry[origin].get('express') or not registry[destination].get('express'):
                row.update(source=tutu.SUGGEST, items=[], statistic='minimum_fare_per_train_and_class',
                           status='unresolved_station')
                return row
            row['source'] = tutu.TRIPS+'?'+urlencode({
                'service':'tutu_trains','term':registry[origin]['express'],
                'term2':registry[destination]['express']})
            row['items'] = tutu.trips(session, registry[origin]['express'], registry[destination]['express'])
            row['statistic'] = 'minimum_fare_per_train_and_class'
            row['status'] = 'ok' if row['items'] else 'no_priced_services'
        elif mode == 'bus':
            row['ambiguous_atlas_endpoint'] = any(registry[n].get('atlas_ambiguous', False)
                                                   for n in (origin, destination))
            row['source'] = ('https://bus.tutu.ru/raspisanie/gorod_'+registry[origin]['bus_slug']
                             +'/gorod_'+registry[destination]['bus_slug']+'/')
            response = session.get(row['source'])
            response.raise_for_status()
            row['items'], no_direct = bus_quotes(response.text,
                registry[origin].get('bus_name', origin), registry[destination].get('bus_name', destination))
            row['statistic'] = 'timetable_indicative_offer_price'
            row['status'] = ('ok' if row['items'] else
                             'no_direct_service' if no_direct else 'no_priced_services')
        else:
            raise ValueError('Unsupported surface mode')
    except tutu.NoPricedServices:
        row.update(status='no_priced_services', reason='schedule_link_only', items=[],
                   statistic='minimum_fare_per_train_and_class')
    except Exception as error:
        response = getattr(error, 'response', None)
        row['status'] = ('source_page_missing' if response is not None
                         and response.status_code == 404 else 'source_error')
        row['error'] = type(error).__name__ + ': ' + str(error).split('?')[0][:180]
        row['items'] = []
    atlas_fallback(session,row)
    row['receipts'] = session.receipts[start:]
    return row


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--raw-dir', type=Path, required=True)
    parser.add_argument('--out', type=Path, default=Path('data/surface/latest.json'))
    parser.add_argument('--resume', action='store_true')
    parser.add_argument('--origins', default='branches', help='branches, all, or comma-separated cities')
    parser.add_argument('--destinations', default='all', help='all or comma-separated cities')
    parser.add_argument('--reverse', action='store_true')
    parser.add_argument('--shard',type=int,default=0)
    parser.add_argument('--of',type=int,default=1)
    args = parser.parse_args()
    if args.of < 1 or not 0 <= args.shard < args.of:
        parser.error('Invalid shard')
    registry = endpoints()
    def choose(value):
        if value == 'all': return list(registry)
        if value == 'branches': return [n for n,c in registry.items() if c['branch_origin']]
        selected = [n.strip() for n in value.split(',')]
        if any(n not in registry for n in selected): parser.error('Unknown registry city')
        return selected
    pairs = select_routes(choose(args.origins),choose(args.destinations),reverse=args.reverse)
    requested = [(a,b,m) for a,b in pairs for m in ('rail','bus')][args.shard::args.of]
    session = ReceiptSession(args.raw_dir)
    snapshot = {'schema_version':2,'started_at':now(),'scope':sorted({n for a,b in pairs for n in (a,b)}),
                'requested_routes':requested,
                'complete':False, 'routes':[]}
    if args.resume and args.out.exists():
        snapshot = json.loads(args.out.read_text(encoding='utf-8'))
        if snapshot['started_at'][:10] != now()[:10]:
            raise ValueError('Resume only within the same UTC collection day')
        if snapshot.get('requested_routes') != [list(v) for v in requested]:
            raise ValueError('Resume must use the identical requested routes/shard')
    existing = {(r['origin'],r['destination'],r['mode']) for r in snapshot['routes']}
    args.out.parent.mkdir(parents=True, exist_ok=True)
    for origin,destination,mode in requested:
        if (origin,destination,mode) in existing: continue
        result = collect_route(session,origin,destination,mode,registry)
        snapshot['routes'].append(result)
        snapshot['updated_at'] = now()
        args.out.write_text(json.dumps(snapshot,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        print(f"{len(snapshot['routes'])}/{len(requested)} {origin} -> {destination} {mode}: "
              f"{result['status']} ({len(result['items'])})",flush=True)
    snapshot['complete'] = len(snapshot['routes']) == len(requested)
    args.out.write_text(json.dumps(snapshot,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')


if __name__ == '__main__':
    main()
