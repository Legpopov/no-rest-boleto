import unittest
from unittest.mock import Mock

from collector.branch_scope import endpoints
from collector.route_scope import route_scope, HUBS
from collector.route_coverage import matrix, summarize
from collector.sources import tutu


class RouteCoverageTests(unittest.TestCase):
    def inputs(self):
        scope = dict(origins=['A'], directed_pairs=[('A', 'B')])
        registry = {'A':dict(iata='AAA'), 'B':dict(iata='BBB')}
        air = dict(source='https://example.org/air', requested=[['AAA','2027-02']],
                   queries=[dict(origin='AAA', month='2027-02', status='ok',
                                 checked_at='2026-11-11', receipts=[])], quotes=[])
        return scope, registry, air, dict(routes=[])

    def test_missing_cache_quote_is_not_no_service_or_zero(self):
        row = next(r for r in matrix(*self.inputs()) if r['requested_period']=='2027-02')
        self.assertEqual(row['status'], 'no_fare')
        self.assertEqual(row['reason'], 'route_not_in_source_cache')
        self.assertEqual(row['reference_price'], '')
        self.assertFalse(row['reference_available'])

    def test_errors_and_unresolved_endpoints_are_not_missing_fares(self):
        scope, registry, air, surface = self.inputs()
        air['queries'][0].update(status='source_error', error_type='HTTPError')
        self.assertEqual(next(r for r in matrix(scope, registry, air, surface) if r['requested_period']=='2027-02')['status'], 'source_error')
        registry['A']['iata'] = ''
        self.assertEqual(next(matrix(scope, registry, air, surface))['status'], 'endpoint_unresolved')

    def test_surface_reference_does_not_create_monthly_observations(self):
        scope, registry, air, surface = self.inputs()
        surface['routes'] = [dict(origin='A',destination='B',mode='rail',status='ok',
                                  source='https://example.org/rail',checked_at='2026-11-11', receipts=[])]
        row = next(r for r in matrix(scope, registry, air, surface) if r['mode']=='rail')
        self.assertEqual(row['status'], 'reference_only')
        self.assertEqual(row['base_period'], '')
        self.assertTrue(row['reference_available'])
        self.assertEqual(summarize(scope, registry, air, surface)['spring_2027_complete'], 0)

    def test_historical_month_keeps_its_actual_source_and_year(self):
        scope, registry, air, surface = self.inputs()
        history=[dict(origin='A',destination='B',basis='one_way',currency='RUB',
                      travel_month='2026-09',price=1234,source='https://example.org/history',
                      retrieved_at='2026-11-10',source_sha256='abc')]
        row=next(r for r in matrix(scope,registry,air,surface,history)
                 if r['mode']=='air' and r['requested_period']=='2026-09')
        self.assertEqual(row['base_period'],'2026-09')
        self.assertEqual(row['source'],'https://example.org/history')
        self.assertEqual(row['reference_price'],1234)
        self.assertEqual(row['query_status'],'not_queried')

    def test_supplement_price_keeps_its_own_receipt_and_check_time(self):
        scope,registry,air,surface=self.inputs()
        air['quotes']=[dict(origin='AAA',destination='BBB',departure_at='2027-02-01',price=1000,
                            source='https://example.org/supplement',retrieved_at='2026-11-11T15:00:00Z',
                            receipt_sha256='supplement-receipt')]
        row=next(r for r in matrix(scope,registry,air,surface)
                 if r['mode']=='air' and r['requested_period']=='2027-02')
        self.assertEqual(row['source'],'https://example.org/supplement')
        self.assertEqual(row['checked_at'],'2026-11-11T15:00:00Z')
        self.assertEqual(row['receipt_sha256'],'supplement-receipt')

    def test_academic_semester_does_not_pool_different_years(self):
        scope, registry, air, surface = self.inputs()
        air['quotes'] = [dict(origin='AAA',destination='BBB',departure_at=f'{year}-{month:02}-01',
                              price=100,source='https://example.org/air')
                         for year, month in [(2027,2),(2027,3),(2027,4),(2026,5),(2026,6)]]
        air['requested'] = [['AAA', q['departure_at'][:7]] for q in air['quotes']]
        air['queries'] = [dict(origin='AAA',month=m,status='ok',checked_at='2026-11-11')
                          for _,m in air['requested']]
        summary = summarize(scope, registry, air, surface)
        self.assertEqual(summary['all_source_spring_month_numbers_complete'], 1)
        self.assertEqual(summary['spring_2027_complete'], 0)

    def test_malformed_railway_response_is_source_error(self):
        session = Mock()
        session.get.return_value.json.return_value = {'error':'Unavailable'}
        with self.assertRaises(ValueError):
            tutu.trips(session, '1', '2')
        session.get.return_value.json.return_value = []
        self.assertEqual(tutu.trips(session, '1', '2'), [])

    def test_scope_contains_every_directed_branch_pair_and_approved_hub(self):
        registry = endpoints()
        if 'Алматы' not in registry:
            registry['Алматы'] = dict(branch_origin=False)
        scope = route_scope(registry)
        self.assertEqual(len(scope['origins']), 50)
        self.assertTrue(set(HUBS) <= set(scope['destinations']))
        pairs = set(scope['directed_pairs'])
        for a in scope['origins']:
            for b in scope['destinations']:
                self.assertEqual((a,b) in pairs, a != b)

    def test_route_link_without_fares_is_distinct_from_malformed_response(self):
        session=Mock()
        session.get.return_value.json.return_value={'url':'/poezda/rasp_d.php?nnst1=1&nnst2=2'}
        with self.assertRaises(tutu.NoPricedServices):
            tutu.trips(session,'1','2')
        with self.assertRaises(ValueError):
            tutu.trips(session,'2','1')


if __name__ == '__main__':
    unittest.main()
