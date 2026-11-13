"""Current ONE-WAY cached reference fares, all available destinations per origin.

The API returns cheapest references, not transaction means. Unknown baggage is
preserved as unknown. No limit on transfer count is imposed by this collector.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re

from .branch_scope import endpoints
from .common import env
from .planning import positive
from .sources.aviasales import URL, months_ahead
from .surface import ReceiptSession, now


def normalize(body, origin, month):
    if body.get('success') is not True or str(body.get('currency','')).upper() != 'RUB':
        raise ValueError('Expected successful explicitly RUB API response')
    result = []
    for item in body['data']:
        if (item.get('origin') != origin or item.get('return_at')
                or not str(item.get('departure_at','')).startswith(month+'-')
                or not positive(item.get('price'))):
            raise ValueError('Air response violates origin/month/one-way/price contract')
        # The source also returns local Cyrillic codes. Keep their raw receipt,
        # but do not interpret them as IATA or discard other valid destinations.
        if not re.fullmatch('[A-Z]{3}',item.get('destination','')):
            continue
        result.append({k:item.get(k) for k in ('origin','destination','price','departure_at',
                       'transfers','airline','flight_number','origin_airport','destination_airport')})
    return result


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--out',type=Path,default=Path('data/air-current/latest.json'))
    parser.add_argument('--raw-dir',type=Path,required=True)
    parser.add_argument('--months',type=int,default=12)
    parser.add_argument('--resume',action='store_true')
    parser.add_argument('--retry-errors',action='store_true')
    parser.add_argument('--shard',type=int,default=0)
    parser.add_argument('--of',type=int,default=1)
    args=parser.parse_args()
    if not 1 <= args.months <= 12 or args.of < 1 or not 0 <= args.shard < args.of:
        parser.error('Invalid horizon or shard')
    registry=endpoints()
    origins=sorted({c['iata'] for c in registry.values() if c['branch_origin']
                    and re.fullmatch('[A-Z]{3}',c.get('iata',''))})
    requested=[(o,m) for o in origins for m in months_ahead(args.months)][args.shard::args.of]
    snapshot=dict(schema_version=1,started_at=now(),requested=[list(v) for v in requested],
                  complete=False,queries=[],quotes=[],
                  statistic='minimum_cached_one_way_reference_per_route_and_month',
                  baggage='unknown_not_reported_by_api',source=URL)
    if args.resume and args.out.exists():
        snapshot=json.loads(args.out.read_text(encoding='utf-8'))
        if snapshot['started_at'][:10]!=now()[:10] or snapshot['requested']!=[list(v) for v in requested]:
            raise ValueError('Resume only identical scope within the same UTC day')
        if args.retry_errors:
            retry={(q['origin'],q['month']) for q in snapshot['queries'] if q['status']=='source_error'}
            snapshot['queries']=[q for q in snapshot['queries'] if (q['origin'],q['month']) not in retry]
            snapshot['quotes']=[q for q in snapshot['quotes'] if (q['origin'],q['departure_at'][:7]) not in retry]
    existing={(q['origin'],q['month']) for q in snapshot['queries']}
    session=ReceiptSession(args.raw_dir)
    token=env('TRAVELPAYOUTS_TOKEN',required=True)
    if 'city_metadata' not in snapshot:
        response=session.get('https://api.travelpayouts.com/data/ru/cities.json')
        response.raise_for_status()
        snapshot['city_metadata']={c['code']:{'name':c['name'],'country_code':c['country_code']}
                                   for c in response.json() if re.fullmatch('[A-Z]{3}',c.get('code',''))}
        snapshot['city_metadata_receipt']=session.receipts[-1]
    args.out.parent.mkdir(parents=True,exist_ok=True)
    for origin,month in requested:
        if (origin,month) in existing:continue
        query=dict(origin=origin,month=month,checked_at=now(),source=URL,status='ok',receipts=[])
        quotes=[];start=len(session.receipts)
        try:
            for page in range(1,21):
                response=session.get(URL,params=dict(origin=origin,departure_at=month,
                    one_way='true',unique='true',direct='false',market='ru',currency='rub',
                    sorting='route',limit=1000,page=page),
                    headers={'X-Access-Token':token,'Accept-Encoding':'gzip, deflate'})
                query['http_status']=response.status_code
                response.raise_for_status()
                body=response.json()
                rows=normalize(body,origin,month)
                query['excluded_local_codes']=query.get('excluded_local_codes',0)+len(body['data'])-len(rows)
                for q in rows:
                    q.update(source=response.url,receipt_sha256=session.receipts[-1]['sha256'])
                quotes.extend(rows)
                if len(body['data'])<1000:break
            else:raise ValueError('Pagination limit reached; do not publish truncated query')
            if not quotes:query['status']='no_cached_fares'
            for q in quotes:
                q.update(retrieved_at=query['checked_at'],baggage_included=None)
                q['id']=hashlib.sha256(json.dumps(q,sort_keys=True).encode()).hexdigest()
        except Exception as error:
            # Never expose request headers, token, or raw exception/URL parameters.
            query['status']='source_error';query['error_type']=type(error).__name__;quotes=[]
            if query.get('http_status') == 400:
                # Store only an allowlisted classification, never a response's
                # arbitrary text. A rejected origin is not a no-service claim.
                try:
                    message=str(response.json().get('error','')).lower()
                    if 'not flightable' in message:
                        query['status']='origin_not_supported'
                except (ValueError, AttributeError):
                    pass
        query['receipts']=session.receipts[start:]
        snapshot['queries'].append(query);snapshot['quotes'].extend(quotes)
        snapshot['updated_at']=now()
        snapshot['complete']=len(snapshot['queries'])==len(requested)
        args.out.write_text(json.dumps(snapshot,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        print(f"{len(snapshot['queries'])}/{len(requested)} {origin} {month}: {query['status']} ({len(quotes)})",flush=True)


if __name__=='__main__':main()
