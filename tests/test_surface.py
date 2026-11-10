import json
import unittest
from collector.surface import bus_quotes, CITIES


def page(offers, currency='RUB'):
    return '<script type="application/ld+json">'+json.dumps({
        '@type':'AggregateOffer','priceCurrency':currency,'offers':offers})+'</script>'


def offer(price=1200, name='Москва – Санкт-Петербург'):
    return {'price':price,'itemOffered':{'name':name,'busNumber':'1',
            'departureTime':'2026-11-08T12:00:00+03:00','provider':{'name':'Provider'}}}


class SurfaceTests(unittest.TestCase):
    def test_directed_registry_has_56_pairs(self):
        self.assertEqual(len(CITIES)*(len(CITIES)-1),56)
        self.assertEqual(CITIES['Казань'][0],'2060615')
        self.assertEqual(CITIES['Минеральные Воды'][1],'Mineralnye-Vody')

    def test_bus_price_is_one_way_and_duplicates_do_not_add_weight(self):
        rows, no_direct = bus_quotes(page([offer(),offer()]),'Москва','Санкт-Петербург')
        self.assertEqual(len(rows),1)
        self.assertEqual(rows[0]['price'],1200)
        self.assertEqual(rows[0]['departure'],'2026-11-08T12:00:00+03:00')
        self.assertFalse(no_direct)

    def test_missing_data_is_not_no_direct_service(self):
        self.assertEqual(bus_quotes('<h1>Расписание</h1>','Москва','Сочи'),([],False))
        self.assertEqual(bus_quotes('<p>Прямых автобусов по этому маршруту нет</p>',
                                   'Москва','Сочи'),([],True))

    def test_wrong_direction_and_currency_cannot_become_estimates(self):
        with self.assertRaisesRegex(ValueError,'direction'):
            bus_quotes(page([offer(name='Сочи – Москва')]),'Москва','Санкт-Петербург')
        with self.assertRaisesRegex(ValueError,'RUB'):
            bus_quotes(page([offer()],currency='USD'),'Москва','Санкт-Петербург')

    def test_invalid_prices_and_conflicts(self):
        self.assertEqual(bus_quotes(page([offer(0),offer(-1)]),'Москва','Санкт-Петербург')[0],[])
        with self.assertRaisesRegex(ValueError,'Conflicting'):
            bus_quotes(page([offer()])+'Прямых автобусов по этому маршруту нет',
                       'Москва','Санкт-Петербург')
