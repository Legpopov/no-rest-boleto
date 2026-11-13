"""Geographic presence and semester-month completeness are different measures."""
from collections import defaultdict


def coverage_report(catalog,registry):
    origins=[n for n,c in registry.items() if c['branch_origin']]
    outgoing=defaultdict(lambda:defaultdict(set))
    air=defaultdict(set)
    for r in catalog['records']:
        outgoing[r['origin']][r['mode']].add(r['destination'])
        if r['mode']=='air' and r['rate_method'].startswith('air_current_'):
            air[(r['origin'],r['destination'])].add(r['month'])
    return dict(retrieved_at=catalog['retrieved_at'],origin_count=len(origins),
                catalog_city_count=len(catalog['cities']),transport_records=len(catalog['records']),
                current_air_pairs=len(air),
                air_spring_months_complete=sum(set(range(2,7))<=v for v in air.values()),
                air_autumn_months_complete=sum({9,10,11,12,1}<=v for v in air.values()),
                origins=[dict(city=n,priced_destinations={m:len(outgoing[n][m]) for m in ('air','rail','bus')}) for n in origins],
                limitations='At least one month/class reference is not complete semester coverage or guaranteed ticket availability. Class/baggage filters are applied in Excel.')
