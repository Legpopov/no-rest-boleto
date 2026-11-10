import unittest
from unittest.mock import patch

from collector.collect import collect_air


class FlightMonthsTests(unittest.TestCase):
    @patch("collector.collect.aviasales.fetch_round_trips")
    def test_month_follows_departure_date_even_when_week_starts_in_previous_year(self, fetch):
        fetch.return_value = [
            {"depart": "2027-01-01", "return": "2027-01-04", "price": 10000,
             "transfers": 0, "return_transfers": 0, "airline": "SU"},
            {"depart": "2027-01-02", "return": "2027-01-05", "price": 20000,
             "transfers": 1, "return_transfers": 0, "airline": "SU"},
        ]
        result = collect_air(None, {"iata": "KZN"}, ["2027-01"])
        self.assertEqual(result["by_week"][0][0], "2026-12-28")
        self.assertEqual(list(result["by_month"]), ["2027-01"])
        self.assertEqual(result["by_month"]["2027-01"]["mean"], 15000)


if __name__ == "__main__":
    unittest.main()
