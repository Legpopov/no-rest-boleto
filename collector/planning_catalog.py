"""Publish only unambiguous, matched reference fares for the offline Excel model."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from .planning import positive


def build_catalog(quotes, changes):
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
    return {
        'schema_version': 1, 'currency': 'RUB', 'basis': 'one_way',
        'statistic': 'minimum_fare_on_reference_date',
        'forecast_assumption': 'Repeat the matched 2025-2026 route/month rate in each future year.',
        'uncertainty': 'Unquantified; these are reference fares, not monthly transaction means.',
        'retrieved_at': max((q['retrieved_at'] for q in quotes), default=None),
        'records': sorted(records, key=lambda r: (r['origin'], r['destination'], r['month'], r['baggage_included'])),
    }


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--research', type=Path, default=Path('data/research'))
    parser.add_argument('--out', type=Path, default=Path('data/planning.json'))
    args = parser.parse_args()
    read = lambda name: json.loads((args.research/name).read_text(encoding='utf-8'))
    catalog = build_catalog(read('ator-quotes.json'), read('ator-route-yoy.json'))
    args.out.write_text(json.dumps(catalog, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(f"Published {len(catalog['records'])} matched route/month/baggage references")


if __name__ == '__main__':
    main()
