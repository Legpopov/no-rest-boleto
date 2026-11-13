"""Publish sourced one-way references and hotel inputs for the offline Excel model."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from statistics import mean
import hashlib

from .planning import positive
from .planning_scope import CITIES
from .branch_scope import endpoints, with_air_destinations
from .planning_indices import category_index, domestic_pair
from .planning_excel import encode_catalog
from .planning_coverage import coverage_report


def surface_records(snapshot, scenario, indices=None, registry=None):
    """Preserve classes/firm flags; never select organization-specific rules here."""
    scope = snapshot.get('scope',list(CITIES))
    expected = set(map(tuple,snapshot['requested_routes'])) if 'requested_routes' in snapshot else {
        (a,b,m) for a in scope for b in scope if a!=b for m in ('rail','bus')}
    actual = [(r['origin'],r['destination'],r['mode']) for r in snapshot['routes']]
    if (not expected or not snapshot.get('complete') or len(actual)!=len(expected) or set(actual)!=expected
            or any(a==b or m not in ('rail','bus') for a,b,m in expected)):
        raise ValueError(f'Surface snapshot must check all {len(expected)} requested mode/direction combinations')
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
        selected_rate = rate
        rate_source = scenario['source']
        rate_method = 'surface_macro_scenario'
        if indices and registry and domestic_pair(route['origin'],route['destination'],registry,indices):
            index=category_index(indices,'passenger_transport')
            selected_rate=index['rate'];rate_source=index['source']
            rate_method='surface_official_category_scenario'
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
                    'price':round(mean(prices),2),'annual_rate':selected_rate,'baggage_included':False,
                    'source':route['source'],'previous_source':rate_source,
                    'reference_date':route['checked_at'][:10],'previous_date':'',
                    'retrieved_at':route['checked_at'],
                    'quote_id':quote_id,'previous_id':'','rate_method':rate_method,
                    'rail_class':cls,'firm':firm,'sample_count':len(prices),
                    'price_method':route['statistic'],
                    'date_method':'retrieval_date_not_travel_month',
                    'reference_departures':sorted({i.get('departure','') for i in route['items']
                                                    if i.get('departure')}),
                    'uncertainty':'Unquantified undated/timetable fare; flat seasonality. Growth repeats a broad sector proxy or macro scenario, not a route forecast.',
                })
    return output


def current_air_records(snapshot, indices, registry, scenario=None):
    """Use actual upcoming travel-month references; never backdate them to 2026."""
    if not snapshot.get('complete'):
        raise ValueError('Incomplete current-air collection')
    requested=set(map(tuple,snapshot['requested']))
    actual=[(q['origin'],q['month']) for q in snapshot['queries']]
    if len(actual)!=len(requested) or set(actual)!=requested:
        raise ValueError('Incomplete/duplicate current-air query scope')
    by_iata={c['iata']:n for n,c in registry.items() if c.get('iata')}
    index=category_index(indices,'passenger_transport')
    groups={}
    for q in snapshot['quotes']:
        a=by_iata.get(q['origin']);b=by_iata.get(q['destination'])
        if not a or not b or a==b:continue
        key=(a,b,q['departure_at'][:7])
        if key not in groups or q['price']<groups[key]['price']:groups[key]=q
    out=[]
    for (a,b,period),q in sorted(groups.items()):
        domestic=domestic_pair(a,b,registry,indices)
        if not domestic and not scenario:continue
        rate=index['rate'] if domestic else scenario['annual_rate']
        rate_source=index['source'] if domestic else scenario['source']
        out.append(dict(origin=a,destination=b,mode='air',month=int(period[5:7]),base_year=int(period[:4]),
            price=q['price'],annual_rate=rate,baggage_included=False,
            source=q['source'],previous_source=rate_source,reference_date=q['departure_at'][:10],
            previous_date='',quote_id=q['id'],previous_id='',rate_method='air_current_official_category_scenario' if domestic else 'air_current_macro_scenario',
            rail_class='',firm=False,sample_count=1,price_method='minimum_cached_one_way_unknown_baggage',
            uncertainty='Unquantified cached minimum; not a sales mean. Baggage unknown; Excel applies its explicit allowance. Growth repeats a sector proxy or RUB macro scenario; FX/local prices are not modeled separately.',
            transfers=q.get('transfers'),retrieved_at=q['retrieved_at']))
    return out


def hotel_quotes(snapshot, indices, registry, scenario):
    out=[]
    index=category_index(indices,'hotels')
    for c in snapshot['cities']:
        h=c.get('hotels',{})
        if h.get('status')!='ok' or h.get('stale') or not h.get('updated_at') or not h.get('items'):continue
        domestic=(c.get('country_code')=='RU' and c['city'] not in indices['excluded_territories'])
        out.append(dict(city=c['city'],abroad=c.get('abroad',True),items=h['items'],
                        retrieved_at=h['updated_at'],source='https://ostrovok.ru/hotel/'+h['slug']+'/',
                        annual_rate=index['rate'] if domestic else scenario['annual_rate'],
                        rate_source=index['source'] if domestic else scenario['source'],
                        rate_method='hotel_official_category_scenario' if domestic else 'hotel_macro_scenario'))
    return out


def build_catalog(quotes, changes, surface=None, scenario=None, current_air=None, indices=None, registry=None, hotels=None):
    if current_air and registry:
        registry=with_air_destinations(registry,current_air)
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
            'retrieved_at':current['retrieved_at'],
        })
    if surface is not None:
        records.extend(surface_records(surface,scenario,indices,registry))
    if current_air is not None:
        records.extend(current_air_records(current_air,indices,registry,scenario))
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
    if current_air:
        current_pairs={(r['origin'],r['destination']) for r in records if r['rate_method'].startswith('air_current_')}
        # Distinguish a missing cache quote from an unqueried/unknown route.
        coverage=[c for c in coverage if c['mode']!='air' or
                  (c['origin'],c['destination']) not in current_pairs]
        for a,b in sorted(current_pairs):
            coverage.append(dict(origin=a,destination=b,mode='air',status='partial_current_reference',
                                 checked_at=current_air['updated_at'],source=current_air['source']))
    return {
        'schema_version': 3, 'currency': 'RUB', 'basis': 'one_way',
        'cities':list(registry or CITIES),'coverage':coverage,
        'city_countries':{n:c.get('country_code','') for n,c in (registry or {}).items()},
        'statistic': 'mixed_reference_fares',
        'forecast_assumption': 'Current domestic one-way air references where available, with a sourced sector-proxy scenario; historical route/month references remain a fallback. Surface references have flat seasonality and an explicit sector-proxy or macro scenario. No implied full geographic coverage.',
        'uncertainty': 'Unquantified; these are reference fares, not monthly transaction means.',
        'surface_scenario':scenario,
        'official_indices':indices,
        'hotel_quotes':hotel_quotes(hotels,indices,registry,scenario) if hotels else [],
        'retrieved_at': max([q['retrieved_at'] for q in quotes]+
                            ([surface['updated_at']] if surface else [])+
                            ([current_air['updated_at']] if current_air else []), default=None),
        'records': sorted(records, key=lambda r: (r['origin'], r['destination'],r['mode'],
                                                 r['month'],r['baggage_included'],
                                                 r.get('rail_class',''),r.get('firm',False))),
    }


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--research', type=Path, default=Path('data/research'))
    parser.add_argument('--out', type=Path, default=Path('data/planning.json'))
    parser.add_argument('--excel-out',type=Path,default=Path('data/planning-excel.json'))
    parser.add_argument('--surface',type=Path,default=Path('data/surface/latest.json'))
    parser.add_argument('--scenario',type=Path,default=Path('data/planning-scenario.json'))
    parser.add_argument('--air-current',type=Path,default=Path('data/air-current/latest.json'))
    parser.add_argument('--indices',type=Path,default=Path('data/planning-indices.json'))
    parser.add_argument('--hotels',type=Path,default=Path('data/planning-hotels.json'))
    args = parser.parse_args()
    read = lambda name: json.loads((args.research/name).read_text(encoding='utf-8'))
    surface = json.loads(args.surface.read_text(encoding='utf-8')) if args.surface.exists() else None
    scenario = json.loads(args.scenario.read_text(encoding='utf-8')) if surface else None
    current_air=json.loads(args.air_current.read_text(encoding='utf-8')) if args.air_current.exists() else None
    indices=json.loads(args.indices.read_text(encoding='utf-8'))
    catalog = build_catalog(read('ator-quotes.json'), read('ator-route-yoy.json'),surface,scenario,
                            current_air,indices,endpoints(),json.loads(args.hotels.read_text(encoding='utf-8')))
    # Expanded coverage repeats thousands of route/month records. Keep one
    # inspectable JSON object per line instead of multiplying whitespace across
    # every field; the document and schema remain ordinary, unchanged JSON.
    metadata = {k:v for k,v in catalog.items() if k!='records'}
    with args.out.open('w',encoding='utf-8') as stream:
        stream.write(json.dumps(metadata,ensure_ascii=False,indent=2)[:-2])
        stream.write(',\n  "records": [\n')
        for position, record in enumerate(catalog['records']):
            if position:
                stream.write(',\n')
            stream.write('    '+json.dumps(record,ensure_ascii=False,separators=(',',':')))
        stream.write('\n  ]\n}\n')
    args.excel_out.write_text(json.dumps(encode_catalog(catalog),ensure_ascii=False,separators=(',',':'))+'\n',encoding='utf-8')
    args.out.with_name('planning-coverage.json').write_text(json.dumps(coverage_report(catalog,endpoints()),ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(f"Published {len(catalog['records'])} transport records, "
          f"{len(catalog['hotel_quotes'])} hotel source sets and compact Excel catalog")


if __name__ == '__main__':
    main()
