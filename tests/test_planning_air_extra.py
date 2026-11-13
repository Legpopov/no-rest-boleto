import unittest
from collector.planning_air_extra import normalize,merge


class SupplementalAirTests(unittest.TestCase):
    def quote(self, **changes):
        return dict(origin='KZN',destination='MOW',depart_date='2027-02-01',
                    return_date='',value=1000,actual=True,trip_class=0,
                    found_at='2026-11-10T10:00:00Z', **changes)

    def test_preserves_exact_period_transfers_and_source_observation_date(self):
        q=self.quote(number_of_changes=2)
        row=normalize(dict(success=True,currency='rub',data=[q]),'KZN','2027-02')[0]
        self.assertEqual(row['departure_at'],'2027-02-01')
        self.assertEqual(row['found_at'],q['found_at'])
        self.assertEqual(row['transfers'],2)
        self.assertIsNone(row['baggage_included'])

    def test_rejects_roundtrip_other_month_currency_and_origin(self):
        for change in ({'return_date':'2027-02-05'},{'depart_date':'2027-03-01'},
                       {'origin':'MOW'},{'value':0}):
            q={**self.quote(),**change}
            with self.assertRaises(ValueError):
                normalize(dict(success=True,currency='rub',data=[q]),'KZN','2027-02')
        with self.assertRaises(ValueError):
            normalize(dict(success=True,currency='usd',data=[]),'KZN','2027-02')

    def test_stale_business_or_unknown_destination_is_not_usable_economy(self):
        for change in ({'actual':False},{'trip_class':1},{'destination':'МАМ'}):
            self.assertEqual(normalize(dict(success=True,currency='rub',data=[{**self.quote(),**change}]),
                                       'KZN','2027-02'),[])

    def test_explicit_empty_response_without_currency_contains_no_fare(self):
        self.assertEqual(normalize(dict(success=True,currency='',data=[]),'KZN','2027-02'),[])
        with self.assertRaises(ValueError):
            normalize(dict(success=True,currency='',data=[self.quote()]),'KZN','2027-02')

    def test_merge_preserves_existing_month_when_new_source_is_empty(self):
        q=dict(origin='KZN',destination='MOW',departure_at='2027-02-01',price=1000,id='original')
        base=dict(quotes=[q],updated_at='2026-11-11T01:00:00Z')
        result=dict(origin='KZN',month='2027-02',status='no_cached_fares',
                    checked_at='2026-11-11T02:00:00Z',quotes=[])
        self.assertEqual(merge(base,[result])['quotes'],[q])
        cheaper={**q,'price':900,'id':'supplement'}
        self.assertEqual(merge(base,[{**result,'quotes':[cheaper]}])['quotes'],[cheaper])


if __name__=='__main__':unittest.main()
