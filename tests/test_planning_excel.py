import json
from pathlib import Path
import unittest
from collector.planning_excel import encode_catalog, decode_records, FIELDS

ROOT=Path(__file__).resolve().parents[1]


class ExcelEncodingTests(unittest.TestCase):
    def test_lossless_all_published_records(self):
        full=json.loads((ROOT/'data/planning.json').read_text(encoding='utf-8'))
        packed=json.loads((ROOT/'data/planning-excel.json').read_text(encoding='utf-8'))
        self.assertEqual(packed,encode_catalog(full))
        decoded=decode_records(packed)
        for original,actual in zip(full['records'],decoded):
            expected={k:original.get(k,1 if k=='sample_count' else False if k in {'firm','baggage_included'} else '') for k in FIELDS}
            self.assertEqual(actual,expected)
        self.assertEqual(len(decoded),len(full['records']))
        self.assertLess((ROOT/'data/planning-excel.json').stat().st_size,(ROOT/'data/planning.json').stat().st_size/3)

    def test_invalid_indices_cannot_become_formula_or_wrong_field(self):
        fixture=json.loads((ROOT/'tests/fixtures/planning-catalog-boundary.json').read_text())
        for bad in [True,-1,99999,'=1+1']:
            packed=encode_catalog(fixture)
            packed['record_values'][0][0]=bad
            with self.assertRaises(ValueError):decode_records(packed)
