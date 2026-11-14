"""Reviewable issue #2 scope, independent of the provider's returned fares."""
from .branch_scope import select_routes

HUBS = ('Новосибирск', 'Сочи', 'Минеральные Воды', 'Владивосток',
        'Хабаровск', 'Иркутск', 'Уфа', 'Омск', 'Челябинск',
        'Минск', 'Ереван', 'Баку', 'Стамбул', 'Дубай', 'Алматы')
INVESTIGATE = ('Донецк', 'Красноармейск', 'Красногорск', 'Ясиноватая')


def route_scope(registry):
    origins = sorted(n for n, c in registry.items() if c['branch_origin'])
    if len(origins) != 50:
        raise ValueError('Expected the reviewed 50 official origins')
    missing = set(HUBS) - registry.keys()
    if missing:
        raise ValueError(f'Unresolved scope cities: {sorted(missing)}')
    destinations = sorted(set(origins) | set(HUBS))
    return dict(origins=origins, destinations=destinations,
                hubs=list(HUBS), directed_pairs=select_routes(origins, destinations),
                policy='Explicit one-way legs; through air fares may include transfers. '
                       'No neighboring airport/station substitution.')
