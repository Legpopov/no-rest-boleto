import copy
import tempfile
import unittest
from pathlib import Path

from collector.history import archive, load_history, observations


def snapshot():
    return {"generated_at": "2026-11-10T05:00:00+00:00", "origin": {"city": "Москва", "iata": "MOW"},
            "cities": [{"id": "c1", "city": "Казань", "country_code": "RU",
                        "air": {"status": "ok", "updated_at": "2026-11-09T05:00:00+00:00",
                                "by_week": [["2026-12-28", 2, 8000, 9000]]}}]}


class HistoryTests(unittest.TestCase):
    def test_replay_and_stale_copy_do_not_create_fresh_observation(self):
        s = snapshot()
        stale = copy.deepcopy(s)
        stale["generated_at"] = "2026-11-10T06:00:00+00:00"
        stale["cities"][0]["air"]["stale"] = True
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(archive([s, stale], d)["added"], 1)
            before = (Path(d) / "2026-11.jsonl.gz").read_bytes()
            self.assertEqual(archive([s], d)["added"], 0)
            self.assertEqual(before, (Path(d) / "2026-11.jsonl.gz").read_bytes())
            row = next(iter(load_history(d).values()))
            self.assertTrue(row["observed_at"].startswith("2026-11-09"))
            self.assertEqual(row["data"]["by_week"][0][0], "2026-12-28")

    def test_origin_is_part_of_transport_identity(self):
        a, b = snapshot(), snapshot()
        b["origin"] = {"city": "Санкт-Петербург", "iata": "LED"}
        self.assertNotEqual(next(observations(a))["id"], next(observations(b))["id"])

    def test_missing_date_is_not_invented(self):
        s = snapshot()
        del s["cities"][0]["air"]["updated_at"]
        self.assertEqual(list(observations(s)), [])

    def test_collection_start_can_precede_block_completion(self):
        s = snapshot()
        s["cities"][0]["air"]["updated_at"] = "2026-11-10T05:11:00+00:00"
        row = next(observations(s))
        self.assertEqual(row["observed_at"], "2026-11-10T05:11:00+00:00")
        self.assertEqual(row["first_seen_at"], row["observed_at"])

    def test_nonfinite_price_is_rejected(self):
        s = snapshot()
        s["cities"][0]["air"]["price"] = float("nan")
        with self.assertRaises(ValueError):
            list(observations(s))


if __name__ == "__main__":
    unittest.main()
