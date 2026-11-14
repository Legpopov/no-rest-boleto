"""Source-backed origin registry; destinations are not restricted to this list."""
from __future__ import annotations

import json
from pathlib import Path
import re

REGISTRY = Path(__file__).resolve().parents[1] / 'data' / 'planning-cities.json'


def city_key(value):
    return re.sub(r'\s+', ' ', value.casefold().replace('ё', 'е').replace('_', ' ')).strip()


def load_registry(path=REGISTRY):
    root = json.loads(Path(path).read_text(encoding='utf-8'))
    cities = root['cities']
    keys = [city_key(c['city']) for c in cities]
    if len(keys) != len(set(keys)):
        raise ValueError('Ambiguous city names in planning registry')
    return root


def endpoints(path=REGISTRY):
    return {c['city']: c for c in load_registry(path)['cities']}


def select_routes(origins, destinations, *, reverse=False):
    """Explicit directed pairs. Never substitute Moscow or mirror a fare."""
    pairs = {(a, b) for a in origins for b in destinations if city_key(a) != city_key(b)}
    if reverse:
        pairs |= {(b, a) for a, b in pairs}
    return sorted(pairs)


def with_air_destinations(registry, snapshot):
    """Resolve only source-returned codes, without inventing nearest airports."""
    result={n:dict(c) for n,c in registry.items()}
    known={c.get('iata') for c in result.values() if c.get('iata')}
    metadata=snapshot.get('city_metadata',{})
    wanted={q['destination'] for q in snapshot['quotes']}
    for code in sorted(wanted-known):
        c=metadata.get(code)
        if not c:continue
        name=c['name']
        if name in result:name=f"{name} ({c['country_code']})"
        if name in result:continue
        result[name]=dict(city=name,country_code=c['country_code'],iata=code,
                          branch_origin=False,express='',bus_slug='',
                          iata_source='https://api.travelpayouts.com/data/ru/cities.json')
    return result
