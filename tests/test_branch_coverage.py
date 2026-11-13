import copy
import json
from pathlib import Path
import unittest

from collector.branch_scope import load_registry, endpoints, select_routes
from collector.planning_air import normalize
from collector.planning_indices import category_index, domestic_pair
from collector.planning_catalog import current_air_records, hotel_quotes, surface_records
from collector.surface import atlas_quotes

BASE=Path(__file__).resolve().parents[1]/'data'


class BranchCoverageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.registry=endpoints()
        cls.indices=json.loads((BASE/'planning-indices.json').read_text(encoding='utf-8'))
        cls.scenario=json.loads((BASE/'planning-scenario.json').read_text(encoding='utf-8'))

    def test_official_origins_are_independent_of_destinations(self):
        origins=[n for n,c in self.registry.items() if c['branch_origin']]
        self.assertEqual(len(origins),50)
        self.assertIn('Воронеж',origins)
        self.assertIn('Ясиноватая',origins)
        pairs=select_routes(['Воронеж'],['Сидней','Москва'],reverse=True)
        self.assertIn(('Воронеж','Сидней'),pairs)
        self.assertIn(('Сидней','Воронеж'),pairs)
        self.assertNotIn(('Москва','Сидней'),pairs)

    def test_no_neighboring_airport_is_silently_substituted(self):
        self.assertEqual(self.registry['Пятигорск']['iata'],'')
        self.assertEqual(self.registry['Выборг']['iata'],'')

    def test_same_month_yoy_is_not_ytd_and_is_geographically_bounded(self):
        self.assertAlmostEqual(category_index(self.indices,'hotels')['rate'],.0587)
        self.assertAlmostEqual(category_index(self.indices,'passenger_transport')['rate'],.1189)
        self.assertFalse(domestic_pair('Москва','Сидней',self.registry,self.indices))
        self.assertFalse(domestic_pair('Москва','Донецк',self.registry,self.indices))
        wrong={**self.indices,'unit':'percent_since_december'}
        with self.assertRaises(ValueError):category_index(wrong,'hotels')

    def test_unknown_local_airport_code_does_not_drop_valid_quote(self):
        row=dict(origin='KZN',destination='LED',departure_at='2027-02-01',price=1000)
        rows=normalize(dict(success=True,currency='rub',data=[row,{**row,'destination':'МАМ'}]),'KZN','2027-02')
        self.assertEqual(len(rows),1)
        self.assertEqual(rows[0]['price'],1000)
        with self.assertRaises(ValueError):normalize(dict(success=True,currency='rub',data=[{**row,'return_at':'2027-02-03'}]),'KZN','2027-02')

    def test_current_air_keeps_actual_base_year_and_unknown_baggage(self):
        q=dict(origin='KZN',destination='LED',departure_at='2027-02-01',price=1000,
               id='test',retrieved_at='2026-11-11T00:00:00Z',source='https://api.travelpayouts.com/',transfers=2)
        s=dict(complete=True,requested=[['KZN','2027-02']],queries=[dict(origin='KZN',month='2027-02')],quotes=[q])
        row=current_air_records(s,self.indices,self.registry,self.scenario)[0]
        self.assertEqual(row['base_year'],2027)
        self.assertFalse(row['baggage_included'])
        self.assertEqual(row['transfers'],2)
        self.assertEqual(row['price'],1000)
        with self.assertRaises(ValueError):current_air_records({**s,'queries':[]},self.indices,self.registry,self.scenario)

    def test_foreign_air_uses_explicit_rub_scenario_not_russian_category(self):
        q=dict(origin='KZN',destination='SYD',departure_at='2027-02-01',price=1000,
               id='test',retrieved_at='2026-11-11T00:00:00Z',source='https://api.travelpayouts.com/')
        s=dict(complete=True,requested=[['KZN','2027-02']],queries=[dict(origin='KZN',month='2027-02')],quotes=[q])
        row=current_air_records(s,self.indices,self.registry,self.scenario)[0]
        self.assertEqual(row['rate_method'],'air_current_macro_scenario')
        self.assertEqual(row['annual_rate'],self.scenario['annual_rate'])

    def test_stale_hotel_cannot_become_current_base(self):
        h=dict(status='ok',stale=True,updated_at='2026-11-11T00:00:00Z',items=[[1000]],slug='russia/moscow')
        c=dict(city='Москва',country_code='RU',abroad=False,hotels=h)
        self.assertEqual(hotel_quotes(dict(cities=[c]),self.indices,self.registry,self.scenario),[])
        h['stale']=False
        self.assertEqual(hotel_quotes(dict(cities=[c]),self.indices,self.registry,self.scenario)[0]['items'],[[1000]])

    def test_dynamic_surface_scope_rejects_duplicates_and_missing_routes(self):
        surface=json.loads((BASE/'surface/latest.json').read_text(encoding='utf-8'))
        row=next(r for r in surface['routes'] if r['status']=='ok')
        small=dict(complete=True,scope=[row['origin'],row['destination']],
                   requested_routes=[[row['origin'],row['destination'],row['mode']]],routes=[row])
        self.assertTrue(surface_records(small,self.scenario))
        with self.assertRaises(ValueError):surface_records({**small,'routes':[row,row]},self.scenario)

    def test_atlas_validates_direction_currency_and_does_not_mix_extrema(self):
        seo=dict(fromName='воронеж',toName='москва',minPrices=[dict(price=1,maxPrice=9999,currency='RUB')])
        day=dict(date='2026-11-20',rideCount=1,minPrices=[dict(price=1000,currency='RUB')])
        root=dict(props=dict(pageProps=dict(seoDirection=seo),initialState=dict(searchRides=dict(calendar=[day]))))
        page=lambda:'<script id="__NEXT_DATA__">'+json.dumps(root)+'</script>'
        self.assertEqual([r['price'] for r in atlas_quotes(page(),'Воронеж','Москва')],[1000])
        with self.assertRaises(ValueError):atlas_quotes(page(),'Москва','Воронеж')
        day['minPrices'][0]['currency']='EUR'
        with self.assertRaises(ValueError):atlas_quotes(page(),'Воронеж','Москва')
