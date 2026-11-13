import unittest
from unittest.mock import patch
from collector.connecting_surface import carrier_quotes,bank_tariffs,community_tariff
from collector.planning_update import merge_surface


class ConnectingSurfaceTests(unittest.TestCase):
    def test_published_tariff_survives_failed_directory_check_with_evidence(self):
        key=dict(origin='A',destination='B',mode='bus')
        good=dict(**key,status='ok',checked_at='2026-11-13T00:00:00Z',
                  statistic='published_official_bank_card_route_tariffs',items=[dict(price=78)])
        failed=dict(**key,status='source_error',checked_at='2026-11-14T00:00:00Z',items=[])
        snapshot=lambda r:dict(routes=[r],started_at=r['checked_at'],updated_at=r['checked_at'])
        with patch('collector.planning_update.surface_records'):
            for sequence in ([good,failed],[failed,good]):
                merged=merge_surface([snapshot(r) for r in sequence],{})
                row=merged['routes'][0]
                self.assertEqual(row['items'][0]['price'],78)
                self.assertEqual(row['checked_at'],good['checked_at'])
                self.assertEqual(row['previous_source_check']['status'],'source_error')
    def test_carrier_uses_directional_table_not_conflicting_marketing_copy(self):
        page='<p>Стоимость 2800 рублей</p><table><tr><th>Рейс</th><th>Станция отправления</th><th>Станция прибытия</th><th>Дни</th><th>Время</th></tr><tr><td>Донецк - Москва 3500 р.</td><td>Крытый рынок</td><td>Москва</td><td>Ежедневно</td><td>09:30</td></tr></table>'
        rows=carrier_quotes(page,'Донецк','Москва')
        self.assertEqual(rows[0]['price'],3500)
        with self.assertRaises(ValueError):carrier_quotes(page,'Москва','Донецк')

    def test_bank_tariff_excludes_discounted_troika_prices(self):
        page='Красногорске метро Тушинская. Тройка: 73 и 85 рублей. Проезд по банковской карте на маршруте 1542 будет стоить 78 рубей, на маршруте 1566 — 90 рублей.'
        self.assertEqual([r['price'] for r in bank_tariffs(page)],[78,90])
        with self.assertRaises(ValueError):bank_tariffs(page.replace('Красногорске','Другой город'))

    def test_secondary_tariff_is_bounded_to_specific_route_and_labelled(self):
        page='Новый маршрут № 106 Б Ясиноватая (АС) – Донецк (ДС Центр). Из Донецка обратно. Стоимость проезда — 120 рублей. Другой маршрут: 90 рублей.'
        row=community_tariff(page)[0]
        self.assertEqual(row['price'],120)
        self.assertIn('not independently confirmed',row['uncertainty'])
        with self.assertRaises(ValueError):community_tariff(page.replace('106 Б','109'))


if __name__=='__main__':unittest.main()
