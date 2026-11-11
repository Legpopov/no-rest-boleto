"""Publish only unambiguous, matched reference fares for the offline Excel model."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from statistics import mean
import hashlib

from .planning import positive
from .planning_scope import CITIES


def surface_records(snapshot, scenario):
    """Preserve classes/firm flags; never select organization-specific rules here."""
    expected = {(a,b,m) for a in CITIES for b in CITIES if a!=b for m in ('rail','bus')}
    actual = [(r['origin'],r['destination'],r['mode']) for r in snapshot['routes']]
    if not snapshot.get('complete') or len(actual)!=112 or set(actual)!=expected:
        raise ValueError('Surface snapshot must check all 112 mode/direction combinations')
    rate = scenario['annual_rate']
    if not isinstance(rate,(int,float)) or isinstance(rate,bool) or not -1 < rate < 1 or not scenario.get('source'):
        raise ValueError('Invalid sourced forecast scenario')
    output = []
    for route in snapshot['routes']:
        if route['status']!='ok': continue
        if route['basis']!='one_way' or route['currency']!='RUB':
            raise ValueError('Surface prices must be one-way RUB')
        if not route.get('receipts') or route['receipts'][-1]['status']!=200:
            raise ValueError('Successful surface route requires a successful HTTP receipt')
        groups, seen = {}, set()
        for item in route['items']:
            encoded = json.dumps(item,ensure_ascii=False,sort_keys=True)
            if encoded in seen: continue
            seen.add(encoded)
            if route['mode']=='rail':
                for cls,price in item['classes'].items():
                    if positive(price):
                        groups.setdefault((cls,bool(item['firm'])),[]).append(price)
            elif positive(item['price']):
                groups.setdefault(('',False),[]).append(item['price'])
        for (cls,firm),prices in sorted(groups.items()):
            quote_id = hashlib.sha256(json.dumps({
                'route':(route['origin'],route['destination'],route['mode']),
                'class':cls,'firm':firm,'prices':prices,'receipts':route['receipts'],
            },sort_keys=True).encode()).hexdigest()
            for month in range(1,13):
                output.append({
                    'origin':route['origin'],'destination':route['destination'],'mode':route['mode'],
                    'month':month,'base_year':int(route['checked_at'][:4]),
                    'price':round(mean(prices),2),'annual_rate':rate,'baggage_included':False,
                    'source':route['source'],'previous_source':scenario['source'],
                    'reference_date':route['checked_at'][:10],'previous_date':'',
                    'quote_id':quote_id,'previous_id':'','rate_method':'surface_macro_scenario',
                    'rail_class':cls,'firm':firm,'sample_count':len(prices),
                    'price_method':route['statistic'],
                    'date_method':'retrieval_date_not_travel_month',
                    'reference_departures':sorted({i.get('departure','') for i in route['items']
                                                    if i.get('departure')}),
                    'uncertainty':'Unquantified: undated/timetable reference, flat seasonality, general inflation scenario.',
                })
    return output


def build_catalog(quotes, changes, surface=None, scenario=None):
    by_id = {q['id']: q for q in quotes}
    records, keys = [], set()
    for change in changes:
        current = by_id[change['after_id']]
        previous = by_id[change['before_id']]
        comparable = ('origin','destination','mode','basis','currency','baggage_included')
        if (any(current[k] != previous[k] for k in comparable)
                or current['travel_month'] != f"2026-{change['month']:02d}"
                or previous['travel_month'] != f"2025-{change['month']:02d}"):
            raise ValueError('Mismatched route/month comparison')
        if (current['basis'] != 'one_way' or current['currency'] != 'RUB'
                or not positive(current['price']) or not positive(previous['price'])):
            raise ValueError('Excel catalog requires positive one-way RUB fares')
        month = change['month']
        key = (current['origin'], current['destination'], current['mode'],
               month, current['baggage_included'])
        if key in keys:
            raise ValueError(f'Ambiguous route/month: {key}')
        keys.add(key)
        records.append({
            'origin': current['origin'], 'destination': current['destination'],
            'mode': current['mode'], 'month': month, 'base_year': change['year'],
            'price': current['price'], 'annual_rate': current['price']/previous['price']-1,
            'baggage_included': current['baggage_included'],
            'source': current['source'], 'previous_source': previous['source'],
            'reference_date': current['travel_date'], 'previous_date': previous['travel_date'],
            'quote_id': current['id'], 'previous_id': previous['id'],
            'rate_method': 'route_reference_yoy_repeated_scenario',
        })
    if surface is not None:
        records.extend(surface_records(surface,scenario))
    coverage = []
    for a in CITIES:
        for b in CITIES:
            if a==b: continue
            coverage.append({'origin':a,'destination':b,'mode':'air',
                             'status':'partial_history' if any(r['origin']==a and r['destination']==b and r['mode']=='air' for r in records) else 'no_historical_reference',
                             'checked_at':max((q['retrieved_at'] for q in quotes),default=''),
                             'source':'https://www.atorus.ru/'})
    if surface:
        coverage.extend({k:r.get(k,'') for k in ('origin','destination','mode','status','checked_at','source')}
                        for r in surface['routes'])
    return {
        'schema_version': 2, 'currency': 'RUB', 'basis': 'one_way',
        'cities':list(CITIES),'coverage':coverage,
        'statistic': 'mixed_reference_fares',
        'forecast_assumption': 'Air: repeat matched 2025-2026 route/month rate. Surface: current reference, flat seasonality, sourced macro scenario; not historical transport inflation.',
        'uncertainty': 'Unquantified; these are reference fares, not monthly transaction means.',
        'surface_scenario':scenario,
        'retrieved_at': max([q['retrieved_at'] for q in quotes]+
                            ([surface['updated_at']] if surface else []), default=None),
        'records': sorted(records, key=lambda r: (r['origin'], r['destination'],r['mode'],
                                                 r['month'],r['baggage_included'],
                                                 r.get('rail_class',''),r.get('firm',False))),
    }


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--research', type=Path, default=Path('data/research'))
    parser.add_argument('--out', type=Path, default=Path('data/planning.json'))
    parser.add_argument('--surface',type=Path,default=Path('data/surface/latest.json'))
    parser.add_argument('--scenario',type=Path,default=Path('data/planning-scenario.json'))
    args = parser.parse_args()
    read = lambda name: json.loads((args.research/name).read_text(encoding='utf-8'))
    surface = json.loads(args.surface.read_text(encoding='utf-8')) if args.surface.exists() else None
    scenario = json.loads(args.scenario.read_text(encoding='utf-8')) if surface else None
    catalog = build_catalog(read('ator-quotes.json'), read('ator-route-yoy.json'),surface,scenario)
    args.out.write_text(json.dumps(catalog, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(f"Published {len(catalog['records'])} planning records (historical air references and explicitly labelled surface scenarios)")


if __name__ == '__main__':
    main()
