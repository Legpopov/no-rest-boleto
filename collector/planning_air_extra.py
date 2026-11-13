"""Supplement the 48-hour date cache with the provider's period cache.

Both methods return cached one-way minimum references, not monthly sales means.
Only actual economy quotes for the exact origin/month are accepted. Preserve the
source's found_at separately from our retrieval time. Never invent historical data.
"""
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import hashlib
import json
from pathlib import Path
import re
import threading

from .common import env
from .planning import positive
from .route_campaign import read, write
from .surface import ReceiptSession, now

URL = 'https://api.travelpayouts.com/aviasales/v3/get_latest_prices'


def normalize(body, origin, month):
    # This endpoint explicitly returns currency="" with a successful empty
    # list. No prices need conversion; classify the actual response as no fare.
    if body.get('success') is True and body.get('data') == [] and body.get('currency') == '':
        return []
    if body.get('success') is not True or str(body.get('currency','')).upper() != 'RUB':
        raise ValueError('Expected successful explicitly RUB response')
    out = []
    for q in body['data']:
        if (q.get('origin') != origin or q.get('return_date')
                or not str(q.get('depart_date','')).startswith(month+'-')
                or not positive(q.get('value'))):
            raise ValueError('Supplement violates origin/month/one-way/price contract')
        if (q.get('actual') is not True or q.get('trip_class') != 0
                or not re.fullmatch('[A-Z]{3}', q.get('destination',''))):
            continue
        out.append(dict(origin=origin, destination=q['destination'], price=q['value'],
                        departure_at=q['depart_date'], transfers=q.get('number_of_changes'),
                        found_at=q.get('found_at'), baggage_included=None, trip_class=0))
    return out


def merge(snapshot, results):
    best = {}
    for q in snapshot['quotes'] + [q for r in results for q in r['quotes']]:
        key = q['origin'], q['destination'], q['departure_at'][:7]
        if key not in best or q['price'] < best[key]['price']:
            best[key] = q
    return {**snapshot, 'quotes':[best[k] for k in sorted(best)],
            'supplement_queries':[{k:v for k,v in r.items() if k!='quotes'} for r in results],
            'supplement_source':URL, 'updated_at':max([snapshot['updated_at']]+[r['checked_at'] for r in results])}


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--state-dir', type=Path, required=True)
    parser.add_argument('--snapshot', type=Path, default=Path('data/air-current/latest.json'))
    parser.add_argument('--workers', type=int, default=4, choices=range(1,5))
    args = parser.parse_args()
    snapshot = read(args.snapshot)
    requested = [(q['origin'],q['month']) for q in snapshot['queries']
                 if q['status'] in ('ok','no_cached_fares')]
    args.state_dir.mkdir(parents=True, exist_ok=True)
    journal = args.state_dir/'queries.jsonl'
    if journal.exists():
        content=journal.read_bytes()
        if content and not content.endswith(b'\n'):
            content=content[:content.rfind(b'\n')+1]
            journal.write_bytes(content)
        results=[json.loads(line) for line in content.decode('utf-8').splitlines()]
    else:
        results=[]
    for result in results:
        if result.get('error_type') != 'ValueError' or not result.get('receipts'):
            continue
        try:
            receipts = result['receipts']
            if all(r['status'] == 200 and normalize(read(args.state_dir/'raw'/(r['sha256']+'.raw')),
                   result['origin'],result['month']) == [] for r in receipts):
                result.update(status='no_cached_fares',quotes=[])
                result.pop('error_type',None)
        except (ValueError, KeyError, FileNotFoundError):
            pass
    done = {(r['origin'],r['month']) for r in results}
    if not done <= set(requested) or len(done) != len(results):
        raise ValueError('Supplement journal does not match query scope')
    local = threading.local()
    token = env('TRAVELPAYOUTS_TOKEN', required=True)
    def collect(key):
        origin, month = key
        if not hasattr(local,'session'):
            local.session = ReceiptSession(args.state_dir/'raw')
        session = local.session
        start = len(session.receipts)
        result = dict(origin=origin,month=month,source=URL,checked_at=now(),status='ok',quotes=[])
        try:
            for page in range(1,21):
                response = session.get(URL,params=dict(origin=origin,beginning_of_period=month+'-01',
                    period_type='month',group_by='directions',one_way='true',currency='rub',market='ru',
                    show_to_affiliates='false',sorting='route',limit=1000,page=page),
                    headers={'X-Access-Token':token,'Accept-Encoding':'gzip, deflate'})
                response.raise_for_status()
                body = response.json()
                for q in normalize(body,origin,month):
                    q.update(source=response.url, receipt_sha256=session.receipts[-1]['sha256'],
                             retrieved_at=result['checked_at'])
                    q['id'] = hashlib.sha256(json.dumps(q,sort_keys=True).encode()).hexdigest()
                    result['quotes'].append(q)
                if len(body['data']) < 1000:
                    break
            else:
                raise ValueError('Truncated supplement pagination')
            if not result['quotes']:
                result['status'] = 'no_cached_fares'
        except Exception as error:
            result.update(status='source_error',error_type=type(error).__name__,quotes=[])
        result['receipts'] = session.receipts[start:]
        return result
    with journal.open('a',encoding='utf-8') as log, ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = [executor.submit(collect,key) for key in requested if key not in done]
        for f in as_completed(futures):
            result = f.result()
            log.write(json.dumps(result,ensure_ascii=False)+'\n');log.flush()
            results.append(result)
            if len(results)%25 == 0:
                print(f'{len(results)}/{len(requested)} supplemental queries',flush=True)
    if {(r['origin'],r['month']) for r in results} != set(requested):
        raise ValueError('Incomplete supplement')
    write(args.snapshot, merge(snapshot, results))
    print(f'Merged {len(results)} query receipts',flush=True)


if __name__ == '__main__':
    main()
