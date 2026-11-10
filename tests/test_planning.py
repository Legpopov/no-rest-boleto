import unittest
from collector.planning import (Calendar, semester_months, annual_index, select_index,
                                forecast_month, forecast_semester, choose_transport, route_total)


class PlanningTests(unittest.TestCase):
    def test_cohort_and_semester_years(self):
        self.assertEqual(semester_months(2027, 1, 1), ['2027-09','2027-10','2027-11','2027-12','2028-01'])
        self.assertEqual(semester_months(2027, 2, 1), semester_months(2027, 1, 3))
        self.assertEqual(semester_months(2027, 1, 2)[0], '2028-02')
        self.assertEqual(len(semester_months(2027, 1, 1, Calendar(9,4,2,6))), 4)

    def test_index_priority_and_source(self):
        peer = {'rate':0.1, 'source':'peer A'}
        official = {'rate':0.2, 'source':'official publication'}
        self.assertEqual(select_index(peers=[peer], official=official)['method'], 'peers')
        self.assertEqual(select_index(route=peer, peers=[official])['method'], 'route')
        self.assertIsNone(select_index(route={'rate':0.1}))
        self.assertIsNone(annual_index(100, 0))

    def test_compound_and_missing_index(self):
        rates = {2027:{'rate':0.1,'source':'A'},2028:{'rate':0.2,'source':'B'}}
        self.assertEqual(forecast_month(100,2026,'2028-01',rates)['price'],132)
        self.assertIsNone(forecast_month(100,2026,'2028-01',{})['price'])
        self.assertIsNone(forecast_month(float('nan'),2026,'2027-01',rates)['price'])

    def test_incomplete_semester_is_not_average_of_available_months(self):
        result = forecast_semester(['2026-09','2026-10'], {9:100}, 2026, {})
        self.assertEqual(result['covered_months'], 1)
        self.assertIsNone(result['price'])

    def test_preference_beats_cheapest_and_missing_mode_is_skipped(self):
        estimates = {'air':{'status':'ok','price':200},'rail':{'status':'ok','price':100}}
        self.assertEqual(choose_transport(estimates,['bus','air','rail'])['mode'],'air')

    def test_route_is_not_doubled_and_gaps_are_visible(self):
        legs = [{'origin':'MOW','destination':'LED','basis':'one_way','status':'ok','price':100},
                {'origin':'LED','destination':'VOZ','basis':'one_way','status':'ok','price':200}]
        self.assertEqual(route_total(legs)['price'],300)
        legs[1]['status']='insufficient_data'
        self.assertIsNone(route_total(legs)['price'])
        legs[0]['basis']='round_trip'
        with self.assertRaises(ValueError): route_total(legs)


if __name__ == '__main__':
    unittest.main()
