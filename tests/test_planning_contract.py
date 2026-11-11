"""The same cases are executed by Excel against the array-only VBA domain core."""
import json
from pathlib import Path
import unittest

from collector.planning import Calendar, choose_transport, forecast_semester, semester_months


class SharedContractTests(unittest.TestCase):
    def test_shared_excel_contract(self):
        cases = json.loads((Path(__file__).parent / 'fixtures/planning-contract.json').read_text())['cases']
        for case in cases:
            with self.subTest(case=case['name']):
                months = semester_months(case['launch'], case['cohort'], case['semester'], Calendar(*case['calendar']))
                self.assertEqual(months, case['months'])
                estimates = {}
                for fare in case['fares']:
                    indices = {m: {year: {'rate': fare['rate'], 'source': 'shared fixture scenario'}
                                   for year in range(2027, int(months[-1][:4]) + 1)}
                               for m in fare['months']}
                    estimates[fare['mode']] = forecast_semester(months, dict.fromkeys(fare['months'], fare['price']), 2026, indices)
                selected = choose_transport(estimates, case['preferences'])
                self.assertEqual(selected['mode'], case['mode'])
                self.assertEqual(selected['price'], case['price'])
