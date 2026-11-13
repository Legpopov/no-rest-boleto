"""Evidence matrix for issue #2; no cached quote is not no transport service."""
import argparse
from collections import Counter, defaultdict
import csv
import gzip
import hashlib
import io
import json
from pathlib import Path

from .branch_scope import endpoints
from .route_scope import route_scope
from .route_campaign import read, write


FIELDS = ('origin', 'destination', 'mode', 'requested_period', 'base_period',
          'status', 'reason', 'checked_at', 'source', 'receipt_sha256',
          'reference_price', 'reference_available', 'transfers', 'reference_kind', 'query_status',
          'supplement_query_status', 'found_at')


def matrix(scope, registry, air, surface, history=()):
    queries = {(q['origin'], q['month']): q for q in air['queries']}
    supplements = {(q['origin'],q['month']):q for q in air.get('supplement_queries',[])}
    quotes = {}
    for q in air['quotes']:
        key = q['origin'], q['destination'], q['departure_at'][:7]
        if key not in quotes or q['price'] < quotes[key]['price']:
            quotes[key] = q
    historical = {}
    for q in history:
        if q.get('basis') != 'one_way' or q.get('currency') != 'RUB':
            continue
        key = q['origin'], q['destination'], q['travel_month']
        if key not in historical or q['price'] < historical[key]['price']:
            historical[key] = q
    periods = sorted({m for _, m in air['requested']} |
                     {f'{year}-{month:02}' for year in (2025, 2026) for month in range(1, 13)})
    routes = {(r['origin'], r['destination'], r['mode']): r for r in surface['routes']}
    for a, b in scope['directed_pairs']:
        for mode in ('air', 'rail', 'bus'):
            for period in periods:
                row = dict.fromkeys(FIELDS, '')
                row.update(origin=a, destination=b, mode=mode, requested_period=period,
                           status='not_checked', reason='route_not_requested', reference_available=False)
                if mode == 'air':
                    origin = registry[a].get('iata', '')
                    destination = registry[b].get('iata', '')
                    query = queries.get((origin, period))
                    quote = quotes.get((origin, destination, period))
                    row.update(source=air['source'], reference_kind='cached_one_way_minimum_unknown_baggage')
                    row['query_status'] = query['status'] if query else 'not_queried'
                    row['supplement_query_status'] = supplements.get((origin,period),{}).get('status','not_queried')
                    if not origin.isascii() or len(origin) != 3:
                        row.update(status='endpoint_unresolved', reason='origin_airport_code_unresolved',
                                   checked_at=registry[a].get('endpoint_checked_at', air.get('city_metadata_receipt', {}).get('retrieved_at','')),
                                   source='https://api.travelpayouts.com/data/ru/cities.json',
                                   receipt_sha256=air.get('city_metadata_receipt',{}).get('sha256',''))
                    elif not destination.isascii() or len(destination) != 3:
                        row.update(status='endpoint_unresolved', reason='destination_airport_code_unresolved',
                                   checked_at=registry[b].get('endpoint_checked_at', air.get('city_metadata_receipt', {}).get('retrieved_at','')),
                                   source='https://api.travelpayouts.com/data/ru/cities.json',
                                   receipt_sha256=air.get('city_metadata_receipt',{}).get('sha256',''))
                    elif query:
                        row.update(checked_at=query['checked_at'])
                        receipts = query.get('receipts', [])
                        row['receipt_sha256'] = '|'.join(r['sha256'] for r in receipts)
                        if quote:
                            row.update(status='priced', reason='', base_period=period,
                                       reference_price=quote['price'], reference_available=True,
                                       transfers=quote.get('transfers', ''), source=quote['source'],
                                       found_at=quote.get('found_at',''),
                                       checked_at=quote.get('retrieved_at',query['checked_at']),
                                       receipt_sha256=quote.get('receipt_sha256',row['receipt_sha256']))
                        elif query['status'] in ('ok', 'no_cached_fares'):
                            row.update(status='no_fare', reason='route_not_in_source_cache')
                        elif query['status'] == 'origin_not_supported':
                            row.update(status='source_unsupported', reason='origin_rejected_by_source')
                        else:
                            row.update(status='source_error', reason=query.get('error_type', query['status']))
                    old = historical.get((a, b, period))
                    if old and not row['reference_available']:
                        row.update(status='priced', reason='', base_period=period,
                                   reference_price=old['price'], reference_available=True,
                                   checked_at=old['retrieved_at'], source=old['source'],
                                   receipt_sha256=old.get('source_sha256', ''),
                                   reference_kind='historical_reference_date_minimum_not_monthly_mean')
                else:
                    route = routes.get((a, b, mode))
                    row['reference_kind'] = 'undated_surface_reference_flat_seasonality_scenario'
                    if route:
                        row.update(checked_at=route['checked_at'], source=route['source'],
                                   query_status=route['status'],
                                   receipt_sha256='|'.join(r['sha256'] for r in route.get('receipts', [])))
                        if route['status'] == 'ok':
                            # No fabricated travel/base month for undated railway/bus fares.
                            row.update(status='reference_only', reason='no_monthly_observation',
                                       reference_available=True)
                        else:
                            row.update(status=route['status'], reason='undated_source_status:'+route['status'])
                yield row


def summarize(scope, registry, air, surface, history=()):
    statuses = Counter()
    by_origin = {a: {m: set() for m in ('air', 'rail', 'bus')} for a in scope['origins']}
    priced_months = defaultdict(set)
    checked = set()
    for r in matrix(scope, registry, air, surface, history):
        statuses[r['status']] += 1
        if r['status'] not in ('not_checked', 'endpoint_unresolved'):
            checked.add(r['origin'])
        if r['reference_available']:
            by_origin[r['origin']][r['mode']].add(r['destination'])
        if r['status'] == 'priced':
            priced_months[(r['origin'], r['destination'])].add(r['base_period'])
    def complete(periods):
        return sum(set(periods) <= months for months in priced_months.values())
    # Preserve the original metric for a like-for-like comparison; explicitly
    # separate it from completeness of a real, continuous academic semester.
    all_months = defaultdict(set)
    for q in air['quotes']:
        all_months[q['origin'], q['destination']].add(int(q['departure_at'][5:7]))
    requested_surface = {(a,b,m) for a,b in scope['directed_pairs'] for m in ('rail','bus')}
    scoped_surface = [r for r in surface['routes'] if (r['origin'],r['destination'],r['mode']) in requested_surface]
    return dict(origin_count=len(scope['origins']), directed_pairs=len(scope['directed_pairs']),
                checked_origins=len(checked), matrix_statuses=dict(sorted(statuses.items())),
                surface_expected_checks=len(requested_surface), surface_completed_checks=len(scoped_surface),
                surface_statuses=dict(sorted(Counter(r['status'] for r in scoped_surface).items())),
                air_origin_month_statuses=dict(sorted(Counter(q['status'] for q in air['queries']).items())),
                air_supplement_statuses=dict(sorted(Counter(q['status'] for q in air.get('supplement_queries',[])).items())),
                scoped_priced_air_pairs=len(priced_months),
                scoped_surface_pairs={m:sum(len(v[m]) for v in by_origin.values()) for m in ('rail', 'bus')},
                spring_2027_complete=complete([f'2027-{m:02}' for m in range(2, 7)]),
                autumn_2026_complete=complete(['2026-09','2026-10','2026-11','2026-12','2027-01']),
                all_source_air_pairs=len(all_months),
                all_source_spring_month_numbers_complete=sum(set(range(2,7)) <= v for v in all_months.values()),
                all_source_autumn_month_numbers_complete=sum({9,10,11,12,1} <= v for v in all_months.values()),
                origins=[dict(city=a, priced_destinations={m:len(v[m]) for m in v}) for a,v in by_origin.items()])


def readable_summary(scope, before, after, air, surface):
    metrics = [
        ('Проверенные наземные записи маршрут/вид', 'surface_completed_checks'),
        ('Авиа: направления с любой ценой, весь источник', 'all_source_air_pairs'),
        ('Авиа: номера месяцев февраль–июнь, весь источник', 'all_source_spring_month_numbers_complete'),
        ('Авиа: номера месяцев сентябрь–январь, весь источник', 'all_source_autumn_month_numbers_complete'),
        ('Авиа: все месяцы весны 2027, согласованная география', 'spring_2027_complete'),
        ('Авиа: все месяцы осени 2026, согласованная география', 'autumn_2026_complete')]
    lines = ['# Покрытие маршрутов: результат issue #2', '',
             f"Авиа проверено до {air['updated_at']}; наземный снимок — до {surface['updated_at']}.", '',
             f"{len(scope['origins'])} городов отправления, {len(scope['destinations'])} назначений, "
             f"{len(scope['directed_pairs'])} направленных пар.", '',
             '| Показатель | Снимок 11.11.2026 | После | Изменение |',
             '|---|---:|---:|---:|']
    for label,key in metrics:
        lines.append(f'| {label} | {before[key]} | {after[key]} | {after[key]-before[key]:+d} |')
    for mode,label in [('rail','Поезд: пары с ориентиром'),('bus','Автобус: пары с ориентиром')]:
        old,new = before['scoped_surface_pairs'][mode],after['scoped_surface_pairs'][mode]
        lines.append(f'| {label}, согласованная география | {old} | {new} | {new-old:+d} |')
    lines += ['', 'Наземные записи включают проверки неразрешённых станций. Они не означают HTTP-запрос тарифа.',
              'Полнота номеров месяцев объединяет разные базовые годы. Полнота конкретного семестра считается отдельно.',
              'Наземные цены остаются недатированными ориентирами с плоской сезонностью, а не помесячной историей.', '',
              '## Проверенные города и оставшиеся пробелы', '',
              '| Город отправления | Авиа: назначения с ценой | Поезд: с ориентиром | Автобус: с ориентиром |',
              '|---|---:|---:|---:|']
    for row in after['origins']:
        counts=row['priced_destinations']
        lines.append(f"| {row['city']} | {counts['air']} | {counts['rail']} | {counts['bus']} |")
    lines += ['', '## Статусы наземных проверок', '', '| Статус | Записей |', '|---|---:|']
    lines.extend(f'| {status} | {count} |' for status,count in after['surface_statuses'].items())
    lines += ['', 'Все пропущенные месяцы и участки перечислены в [matrix.csv.gz](matrix.csv.gz), вместе с источниками, датами и причинами.',
              'Точная география находится в [scope.json](scope.json), машинное сравнение — в [summary.json](summary.json).',
              'Расшифровка статусов и разбор четырёх проблемных городов: [методика](../../docs/route-coverage.md).',
              'Пустой тариф не равен нулю. Отсутствие цены, ошибка источника и отсутствие прямого сообщения различаются.', '']
    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument('--baseline-dir', type=Path,
                        help='First run only: frozen pre-campaign snapshots. Later runs reuse the published baseline.')
    parser.add_argument('--out', type=Path, default=Path('data/route-coverage'))
    args = parser.parse_args()
    registry = endpoints()
    scope = route_scope(registry)
    air = read('data/air-current/latest.json')
    surface = read('data/surface/latest.json')
    history = read('data/research/ator-quotes.json')
    baseline_path = args.out/'baseline-2026-11-11.json'
    if args.baseline_dir:
        before_air = read(args.baseline_dir/'baseline-air-current-latest.json')
        before_surface = read(args.baseline_dir/'baseline-surface-latest.json')
        before = summarize(scope, registry, before_air, before_surface, history)
        baseline = dict(air_updated_at=before_air['updated_at'], surface_updated_at=before_surface['updated_at'],
                        scope_sha256=hashlib.sha256(json.dumps(scope, ensure_ascii=False, sort_keys=True).encode()).hexdigest(),
                        sha256={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in args.baseline_dir.glob('baseline-*.json')},
                        summary=before)
        if baseline_path.exists():
            frozen = read(baseline_path)
            if (any(frozen.get(k) != baseline[k] for k in ('sha256','scope_sha256','air_updated_at','surface_updated_at'))
                    or any(baseline['summary'].get(k) != v for k,v in frozen['summary'].items())):
                raise ValueError('Refusing to overwrite the frozen 11 November baseline')
    else:
        baseline = read(baseline_path)
        before = baseline['summary']
        if baseline['scope_sha256'] != hashlib.sha256(json.dumps(scope, ensure_ascii=False, sort_keys=True).encode()).hexdigest():
            raise ValueError('Baseline uses another scope; cannot compare silently')
    after = summarize(scope, registry, air, surface, history)
    args.out.mkdir(parents=True, exist_ok=True)
    write(args.out/'scope.json', scope)
    write(baseline_path, baseline)
    write(args.out/'summary.json', dict(air_updated_at=air['updated_at'], surface_updated_at=surface['updated_at'],
                                      before=before, after=after,
                                      limitations='No price is not no service. Undated surface references are not monthly observations. '
                                      'Month-number completeness pools base years and is not a continuous semester. '
                                      'Only explicit no_direct_service confirms the source reports no direct bus. '
                                      'Unresolved endpoints are not checked routes; see endpoint-investigation.json.'))
    (args.out/'README.md').write_text(readable_summary(scope,before,after,air,surface),encoding='utf-8')
    with (args.out/'matrix.csv.gz').open('wb') as raw, gzip.GzipFile(fileobj=raw, mode='wb', mtime=0) as zipped:
        with io.TextIOWrapper(zipped, encoding='utf-8', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=FIELDS)
            writer.writeheader()
            writer.writerows(matrix(scope, registry, air, surface, history))
    print(json.dumps({k:v for k,v in after.items() if k!='origins'}, ensure_ascii=False), flush=True)


if __name__ == '__main__':
    main()
