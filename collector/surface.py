"""Independent one-way rail/bus snapshots for the eight-city planning scope.

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
from urllib.parse import urlencode

from .common import PoliteSession
from .planning import positive
from .sources import tutu

CITIES = {
    'Москва': ('2000000', 'Moskva'),
    'Санкт-Петербург': ('2004000', 'Sankt-Peterburg'),
    'Екатеринбург': ('2030000', 'Ekaterinburg'),
    'Казань': ('2060615', 'Kazan'),
    'Новосибирск': ('2044000', 'Novosibirsk'),
    'Сочи': ('2064130', 'Sochi'),
    'Минеральные Воды': ('2064170', 'Mineralnye-Vody'),
    'Калининград': ('2058000', 'Kaliningrad'),
}


def now():
    return datetime.now(timezone.utc).isoformat()


class ReceiptSession(PoliteSession):
    def __init__(self, raw_dir):
        super().__init__(1.0, 1.8)
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


def collect_route(session, origin, destination, mode):
    start = len(session.receipts)
    row = {'origin': origin, 'destination': destination, 'mode': mode,
           'currency': 'RUB', 'basis': 'one_way', 'checked_at': now(),
           'price_period': 'undated_reference_not_monthly_history'}
    try:
        if mode == 'rail':
            row['source'] = tutu.TRIPS+'?'+urlencode({
                'service':'tutu_trains','term':CITIES[origin][0],
                'term2':CITIES[destination][0]})
            row['items'] = tutu.trips(session, CITIES[origin][0], CITIES[destination][0])
            row['statistic'] = 'minimum_fare_per_train_and_class'
            row['status'] = 'ok' if row['items'] else 'no_priced_services'
        elif mode == 'bus':
            row['source'] = ('https://bus.tutu.ru/raspisanie/gorod_'+CITIES[origin][1]
                             +'/gorod_'+CITIES[destination][1]+'/')
            response = session.get(row['source'])
            response.raise_for_status()
            row['items'], no_direct = bus_quotes(response.text, origin, destination)
            row['statistic'] = 'timetable_indicative_offer_price'
            row['status'] = ('ok' if row['items'] else
                             'no_direct_service' if no_direct else 'no_priced_services')
        else:
            raise ValueError('Unsupported surface mode')
    except Exception as error:
        response = getattr(error, 'response', None)
        row['status'] = ('source_page_missing' if response is not None
                         and response.status_code == 404 else 'source_error')
        row['error'] = type(error).__name__ + ': ' + str(error).split('?')[0][:180]
        row['items'] = []
    row['receipts'] = session.receipts[start:]
    return row


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--raw-dir', type=Path, required=True)
    parser.add_argument('--out', type=Path, default=Path('data/surface/latest.json'))
    parser.add_argument('--resume', action='store_true')
    args = parser.parse_args()
    session = ReceiptSession(args.raw_dir)
    snapshot = {'schema_version':1,'started_at':now(),'scope':list(CITIES),
                'complete':False, 'routes':[]}
    if args.resume and args.out.exists():
        snapshot = json.loads(args.out.read_text(encoding='utf-8'))
        if snapshot['started_at'][:10] != now()[:10]:
            raise ValueError('Resume only within the same UTC collection day')
    existing = {(r['origin'],r['destination'],r['mode']) for r in snapshot['routes']}
    args.out.parent.mkdir(parents=True, exist_ok=True)
    for origin in CITIES:
        for destination in CITIES:
            if origin == destination: continue
            for mode in ('rail','bus'):
                if (origin,destination,mode) in existing: continue
                result = collect_route(session,origin,destination,mode)
                snapshot['routes'].append(result)
                snapshot['updated_at'] = now()
                args.out.write_text(json.dumps(snapshot,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
                print(f"{len(snapshot['routes'])}/112 {origin} -> {destination} {mode}: "
                      f"{result['status']} ({len(result['items'])})",flush=True)
    snapshot['complete'] = len(snapshot['routes']) == 112
    args.out.write_text(json.dumps(snapshot,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')


if __name__ == '__main__':
    main()
