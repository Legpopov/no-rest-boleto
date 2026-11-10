import json
from pathlib import Path
import unittest

from collector.planning_catalog import build_catalog


class CatalogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        base = Path(__file__).resolve().parents[1] / 'data'
        cls.quotes = json.loads((base/'research/ator-quotes.json').read_text(encoding='utf-8'))
        cls.changes = json.loads((base/'research/ator-route-yoy.json').read_text(encoding='utf-8'))

    def test_published_catalog_reproduces_sources(self):
        actual = build_catalog(self.quotes, self.changes)
        published = Path(__file__).resolve().parents[1]/'data/planning.json'
        self.assertEqual(actual, json.loads(published.read_text(encoding='utf-8')))
        self.assertEqual(len(actual['records']), 188)

    def test_moscow_petersburg_baggage_reference(self):
        records = build_catalog(self.quotes, self.changes)['records']
        row = next(r for r in records if r['origin'] == 'Москва'
                   and r['destination'] == 'Санкт-Петербург'
                   and r['month'] == 5 and r['baggage_included'])
        self.assertEqual(row['price'], 7498)
        self.assertAlmostEqual(row['annual_rate'], 7498/6316-1)
        self.assertEqual(row['rate_method'], 'route_reference_yoy_repeated_scenario')

    def test_duplicate_month_does_not_become_first_match(self):
        with self.assertRaisesRegex(ValueError, 'Ambiguous'):
            build_catalog(self.quotes, self.changes + self.changes[:1])

    def test_roundtrip_is_never_halved(self):
        quotes = [dict(q, basis='round_trip') for q in self.quotes]
        with self.assertRaisesRegex(ValueError, 'one-way'):
            build_catalog(quotes, self.changes)
