import json
from pathlib import Path
import unittest
from collector.routing import best_route,SemesterRouter


class RoutingTests(unittest.TestCase):
    def test_shared_routing_contract(self):
        root=Path(__file__).parent/'fixtures/routing-contract.json'
        for case in json.loads(root.read_text())['cases']:
            with self.subTest(case['name']):
                edges=[dict(zip(('origin','destination','mode','price','transfers'),v)) for v in case['edges']]
                result=best_route(edges,case['origin'],case['destination'],case['preferences'])
                self.assertEqual(result['status'],case['status'])
                if case['status']=='ok':
                    self.assertEqual(result['cities'],case['cities'])
                    self.assertEqual(result['price'],case['price'])
                    self.assertEqual(result['known_transfers'],case['known_transfers'])
                    self.assertEqual(result['transfer_limit_verified'],case['verified'])
                    if len(result['legs'])>1 or not case['verified'] or case['known_transfers']:
                        self.assertTrue(result['ticket_sum_only'])
                        self.assertTrue(result['warnings'])
                else:
                    self.assertIsNone(result['price'])

    def records(self):
        return [dict(origin=a,destination=b,mode='rail',month=m,base_year=2026,price=p,
                     annual_rate=.1,baggage_included=False,rail_class='coupe',firm=True,
                     sample_count=1,source='https://example.org/fare',rate_method='surface_macro_scenario')
                for a,b,p in [('A','B',100),('B','D',200)] for m in range(2,7)]

    def test_all_months_are_required_on_every_leg(self):
        rows=self.records()
        router=SemesterRouter(dict(records=rows))
        result=router.plan('A','D',2026,1,2,['rail'])
        self.assertEqual(result['price'],330)
        self.assertEqual(result['cities'],['A','B','D'])
        # A single missing month in the second leg invalidates the whole path.
        router=SemesterRouter(dict(records=rows[:-1]))
        self.assertEqual(router.plan('A','D',2026,1,2,['rail'])['status'],'insufficient_data')

    def test_class_filter_is_not_bypassed_to_connect_route(self):
        rows=self.records()
        for row in rows:
            if row['origin']=='B':row['firm']=False
        router=SemesterRouter(dict(records=rows))
        self.assertIsNone(router.plan('A','D',2026,1,2,['rail'],rail_classes=('coupe:firm',))['price'])
        self.assertEqual(router.plan('A','D',2026,1,2,['rail'],rail_classes=('coupe',))['price'],330)

    def test_unknown_baggage_requires_explicit_allowance_per_ticket(self):
        rows=[dict(r,mode='air',rate_method='air_current_macro_scenario',transfers=0,baggage_included=False)
              for r in self.records()]
        router=SemesterRouter(dict(records=rows,city_countries={'A':'RU','B':'RU','D':'RU'}))
        self.assertIsNone(router.plan('A','D',2026,1,2,['air'],baggage=True)['price'])
        self.assertEqual(router.plan('A','D',2026,1,2,['air'],baggage=True,allowance_ru=50)['price'],440)

    def test_future_base_year_cannot_price_earlier_semester(self):
        router=SemesterRouter(dict(records=[dict(r,base_year=2029) for r in self.records()]))
        self.assertIsNone(router.plan('A','D',2026,1,2,['rail'])['price'])

    def test_largest_monthly_internal_count_is_used(self):
        rows=[dict(r,mode='air',rate_method='air_current_macro_scenario',transfers=0,baggage_included=False)
              for r in self.records()]
        rows[0]['transfers']=2
        router=SemesterRouter(dict(records=rows))
        self.assertIsNone(router.plan('A','D',2026,1,2,['air'])['price'])


if __name__=='__main__':unittest.main()
