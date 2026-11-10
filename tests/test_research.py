import unittest
from collector.research import route_quotes, pooled_indices, reference_changes

SOURCE={'id':'test','url':'https://example.org/report','sha256':'test','retrieved_at':'2026-11-11'}
HTML='''<p>Минимальная стоимость билетов из Москвы в крупнейшие аэропорты России на 4 мая 2026 года:</p>
<table><tr><td></td><td>С багажом</td><td>Без багажа</td></tr>
<tr><td>в Санкт-Петербург (3 мая)</td><td>7 &#8288;498 руб. Carrier B</td><td>4 876 руб. Carrier A</td></tr></table>'''

class ResearchTests(unittest.TestCase):
    def test_headers_not_positions_and_exception_date(self):
        rows, rejected=route_quotes(HTML,SOURCE)
        self.assertFalse(rejected)
        self.assertEqual({r['baggage_included']:r['price'] for r in rows},{False:4876,True:7498})
        self.assertEqual(rows[0]['travel_date'],'2026-05-03')
        self.assertEqual(rows[0]['statistic'],'minimum_fare_on_reference_date')

    def test_repeat_is_idempotent(self):
        self.assertEqual(route_quotes(HTML,SOURCE),route_quotes(HTML,SOURCE))

    def test_caption_variant_does_not_inherit_previous_origin(self):
        other=HTML.replace('Минимальная стоимость билетов из Москвы','Стоимость билетов из Казани')
        rows,_=route_quotes(HTML+other,SOURCE)
        self.assertEqual([r['origin'] for r in rows],['Москва','Москва','Казань','Казань'])

    def test_short_caption_is_bound_to_its_own_origin(self):
        other=HTML.replace('Минимальная стоимость билетов из Москвы','Из Калининграда')
        rows,_=route_quotes(HTML+other,SOURCE)
        self.assertEqual(rows[-1]['origin'],'Калининград')

    def test_city_case_and_reference_month_exception(self):
        other=HTML.replace('Санкт-Петербург (3 мая)','Минеральные воды (2 апреля)')
        rows,_=route_quotes(other,SOURCE)
        self.assertEqual(rows[0]['destination'],'Минеральные Воды')
        self.assertEqual(rows[0]['travel_month'],'2026-04')

    def test_malformed_currency_not_silently_zero(self):
        with self.assertRaises(ValueError): route_quotes(HTML.replace('руб.','USD'),SOURCE)

    def test_summary_without_route_caption_is_not_a_route(self):
        self.assertEqual(route_quotes(HTML[HTML.index('<table>'):],SOURCE),([],[]))

    def test_pooled_index_keeps_scope_and_decimal_units(self):
        html='<table><tr><td>Категория</td><td>2025 год (руб.)</td><td>2026 год (руб.)</td></tr><tr><td>С багажом (Средняя)</td><td>10 000</td><td>9 000</td></tr></table>'
        row=pooled_indices(html,{**SOURCE,'id':'ator-may-2026'})[0]
        self.assertAlmostEqual(row['rate'],-0.1)
        self.assertEqual(row['scope'],'publisher_route_sample')

    def test_month_change_does_not_mix_baggage_and_keeps_source_ids(self):
        a,_=route_quotes(HTML,SOURCE)
        b,_=route_quotes(HTML.replace('4 мая','1 июня').replace('(3 мая)','(1 июня)'),SOURCE)
        rows=reference_changes(a+b)
        self.assertEqual(len(rows),2)
        self.assertTrue(all(r['rate']==0 for r in rows))
        with self.assertRaises(ValueError): reference_changes(a+a+b)
