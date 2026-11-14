import json
from pathlib import Path
import unittest

from collector.planning_catalog import build_catalog, surface_records
from collector.branch_scope import endpoints
from statistics import mean


class CatalogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        base = Path(__file__).resolve().parents[1] / 'data'
        cls.quotes = json.loads((base/'research/ator-quotes.json').read_text(encoding='utf-8'))
        cls.changes = json.loads((base/'research/ator-route-yoy.json').read_text(encoding='utf-8'))
        cls.surface = json.loads((base/'surface/latest.json').read_text(encoding='utf-8'))
        cls.scenario = json.loads((base/'planning-scenario.json').read_text(encoding='utf-8'))
        cls.indices = json.loads((base/'planning-indices.json').read_text(encoding='utf-8'))
        cls.current = json.loads((base/'air-current/latest.json').read_text(encoding='utf-8'))
        cls.hotels = json.loads((base/'planning-hotels.json').read_text(encoding='utf-8'))

    def test_published_catalog_reproduces_sources(self):
        actual = build_catalog(self.quotes, self.changes, self.surface, self.scenario,
                               self.current,self.indices,endpoints(),self.hotels)
        published = Path(__file__).resolve().parents[1]/'data/planning.json'
        self.assertEqual(actual, json.loads(published.read_text(encoding='utf-8')))
        self.assertEqual(len([r for r in actual['records'] if r['rate_method']=='route_reference_yoy_repeated_scenario']), 881)
        self.assertEqual(sorted({r['month'] for r in actual['records'] if r['rate_method']=='route_reference_yoy_repeated_scenario'}),
                         [3,5,6,7,8,9,10,11,12])
        self.assertEqual({r['mode'] for r in actual['records']},{'air','rail','bus'})

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

    def test_cross_route_join_is_rejected(self):
        wrong = dict(self.changes[0], before_id=self.changes[-1]['before_id'])
        with self.assertRaisesRegex(ValueError, 'Mismatched'):
            build_catalog(self.quotes, [wrong])

    def test_coverage_keeps_empty_months_and_ambiguities(self):
        path = Path(__file__).resolve().parents[1]/'data/research/coverage.json'
        coverage = json.loads(path.read_text(encoding='utf-8'))
        self.assertEqual(len(coverage['months']),24)
        months = {m['month']:m for m in coverage['months']}
        self.assertEqual(months['2026-04']['quotes'],0)
        self.assertGreater(months['2026-01']['quotes'],0)
        self.assertEqual(months['2026-01']['matched_yoy'],0)
        self.assertGreater(coverage['comparison_exclusions'],0)

    def test_partial_surface_collection_is_not_published(self):
        with self.assertRaisesRegex(ValueError,'requested'):
            surface_records({**self.surface,'routes':self.surface['routes'][:-1]},self.scenario)

    def test_surface_mean_is_weighted_by_samples_and_keeps_class_rules(self):
        rows = surface_records(self.surface,self.scenario)
        matched = [r for r in rows if r['origin']=='Москва' and
                   r['destination']=='Санкт-Петербург' and r['mode']=='rail'
                   and r['month']==1 and r['rail_class']=='coupe']
        weighted = sum(r['price']*r['sample_count'] for r in matched)/sum(r['sample_count'] for r in matched)
        route = next(r for r in self.surface['routes'] if r['origin']=='Москва' and
                     r['destination']=='Санкт-Петербург' and r['mode']=='rail')
        expected = mean(t['classes']['coupe'] for t in route['items'] if 'coupe' in t['classes'])
        self.assertAlmostEqual(weighted,expected,delta=.01)
        self.assertEqual({r['firm'] for r in matched},{True,False})
        self.assertTrue(all(r['date_method']=='retrieval_date_not_travel_month' for r in rows))
        self.assertTrue(all(r['rate_method']=='surface_macro_scenario' for r in rows))

    def test_no_direct_bus_is_not_replaced_by_train_or_reverse_fare(self):
        common=dict(currency='RUB',basis='one_way',checked_at='2026-11-11T00:00:00Z',
                    source='https://example.org/fixed-test-reference',
                    statistic='test_reference',receipts=[dict(status=200,sha256='test')])
        snapshot=dict(complete=True,scope=['A','B'],
            requested_routes=[['A','B','bus'],['A','B','rail'],['B','A','bus']],routes=[
                dict(common,origin='A',destination='B',mode='bus',status='no_direct_service',items=[]),
                dict(common,origin='A',destination='B',mode='rail',status='ok',
                     items=[dict(classes={'coupe':2000},firm=False)]),
                dict(common,origin='B',destination='A',mode='bus',status='ok',items=[dict(price=1000)])])
        rows = surface_records(snapshot,self.scenario)
        self.assertFalse(any(r['origin']=='A' and r['destination']=='B' and r['mode']=='bus' for r in rows))
        self.assertTrue(any(r['origin']=='A' and r['destination']=='B' and r['mode']=='rail' for r in rows))
        bus = next(r for r in rows if r['origin']=='B' and r['destination']=='A' and r['mode']=='bus')
        self.assertEqual(bus['price'],1000)
        self.assertEqual(bus['annual_rate'],.04)
        self.assertIn('cbr.ru',bus['previous_source'])
