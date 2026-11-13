"""Merge checked surface scopes and freeze a dated hotel input for reproducibility."""
from __future__ import annotations
import argparse
import json
from pathlib import Path

from .planning_catalog import surface_records


def merge_surface(snapshots, scenario):
    rows={}
    for snapshot in snapshots:
        surface_records(snapshot,scenario)  # Validate scope, duplicates and receipts first.
        for row in snapshot['routes']:
            key=(row['origin'],row['destination'],row['mode'])
            if key not in rows or row['checked_at'] >= rows[key]['checked_at']:rows[key]=row
    return dict(schema_version=2,complete=True,
                started_at=min(s['started_at'] for s in snapshots),
                updated_at=max(s['updated_at'] for s in snapshots),
                scope=sorted({n for a,b,m in rows for n in (a,b)}),
                requested_routes=[list(k) for k in sorted(rows)],routes=[rows[k] for k in sorted(rows)])


def hotel_snapshot(prices):
    return dict(generated_at=prices['generated_at'],cities=[
        {**{k:c[k] for k in ('city','country_code','abroad')},
         'hotels':{k:v for k,v in c.get('hotels',{}).items()
                   if k in ('status','stale','updated_at','items','slug')}}
        for c in prices['cities']])


def main():
    parser=argparse.ArgumentParser(__doc__)
    parser.add_argument('--surface',type=Path,nargs='+',required=True)
    parser.add_argument('--latest',type=Path,default=Path('data/surface/latest.json'))
    parser.add_argument('--hotel-prices',type=Path,default=Path('data/prices.json'))
    parser.add_argument('--hotel-out',type=Path,default=Path('data/planning-hotels.json'))
    args=parser.parse_args()
    read=lambda p:json.loads(p.read_text(encoding='utf-8'))
    snapshots=([read(args.latest)] if args.latest.exists() else [])+[read(p) for p in args.surface]
    merged=merge_surface(snapshots,read(Path('data/planning-scenario.json')))
    args.latest.write_text(json.dumps(merged,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    args.hotel_out.write_text(json.dumps(hotel_snapshot(read(args.hotel_prices)),ensure_ascii=False,indent=2)+'\n',encoding='utf-8')


if __name__=='__main__':main()
