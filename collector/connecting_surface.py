"""Published reference fares for missing city connectors; not live ticket search.

Carrier timetable tables take precedence over marketing copy. Official bank-card
tariffs exclude discount/subscription fares. A separately labelled community
reference is permitted only for the specifically identified Yasynuvata route.
"""
import argparse
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
import re

from .surface import ReceiptSession, now
from .route_campaign import write

SOURCES = {
    'donetsk':'https://autoflotbus.ru/avtobus-doneck-moskva',
    'moscow':'https://autoflotbus.ru/moskva-doneck',
    'krasnogorsk':'https://transport.mos.ru/mostrans/all_news/128678',
    'yasinovataya':'https://ok.ru/group/54563669540864/topic/165683371116544',
}


def text(html):
    return re.sub(r'\s+',' ',unescape(re.sub(r'<[^>]*>',' ',html))).strip()


class Tables(HTMLParser):
    def __init__(self):
        super().__init__();self.rows=[];self.row=None;self.cell=None
    def handle_starttag(self,tag,attrs):
        if tag=='tr':self.row=[]
        if tag in ('td','th') and self.row is not None:self.cell=[]
    def handle_data(self,value):
        if self.cell is not None:self.cell.append(value)
    def handle_endtag(self,tag):
        if tag in ('td','th') and self.cell is not None:
            self.row.append(text(' '.join(self.cell)));self.cell=None
        if tag=='tr' and self.row is not None:self.rows.append(self.row);self.row=None


def carrier_quotes(html,origin,destination):
    parsed=Tables();parsed.feed(html)
    expected=re.sub(r'\s+','',origin+'-'+destination).casefold()
    rows=[];valid_header=False
    for row in parsed.rows:
        if 'Рейс' in row and 'Станция отправления' in row and 'Станция прибытия' in row:
            valid_header=True;continue
        if not valid_header or len(row)<5:continue
        match=re.fullmatch(r'(.+?)\s+(\d[\d ]*)\s*р\.',row[0])
        if not match:continue
        if re.sub(r'\s+','',match[1]).casefold()!=expected:
            raise ValueError('Carrier timetable direction mismatch')
        price=int(match[2].replace(' ',''))
        if price<=0:raise ValueError('Nonpositive carrier fare')
        rows.append(dict(price=price,provider='АвтоФлот',departure_stop=row[1],arrival_stop=row[2],
                         operating_days=row[3],departure=row[4],reference_kind='published_carrier_timetable'))
    if not rows:raise ValueError('Carrier timetable with directional prices not found')
    return rows


def bank_tariffs(html):
    body=text(html)
    if not all(v in body for v in ('Красногорске','Тушинская','1542','1566')):
        raise ValueError('Official tariff geography/route contract changed')
    match=re.search(r'Проезд по банковской карте на маршруте 1542 будет стоить\s+(\d+)\s+руб\w*,\s*на маршруте 1566\s*[—–-]\s*(\d+)\s+руб',body)
    if not match:raise ValueError('Explicit official bank-card tariff sentence not found')
    return [dict(price=int(price),provider='Мосгортранс',number=number,payment='bank_card',
                 tariff_effective_from='2026-03-14',reference_kind='published_official_route_tariff')
            for number,price in zip(('1542','1566'),match.groups())]


def community_tariff(html):
    body=text(html)
    section=re.search(r'Новый маршрут\s*№\s*106\s*Б(.{0,1600}?)Стоимость проезда\s*[—–:-]\s*(\d+)\s+руб',body)
    if not section or not all(v in section[1] for v in ('Ясиноватая','Донецк','Центр','Из Донецка')):
        raise ValueError('Community route identity/directions not established')
    price=int(section[2])
    if not 0<price<10000:raise ValueError('Invalid local community fare')
    return [dict(price=price,provider='Опубликованный местный ориентир',number='106 Б',
                 tariff_effective_from='2026-07-06',reference_kind='secondary_community_tariff_reference',
                 uncertainty='Community publication; price not independently confirmed by operator')]


def collect(raw_dir):
    session=ReceiptSession(raw_dir);started=now();routes=[]
    definitions=[('donetsk','Донецк','Москва'),('moscow','Москва','Донецк'),
                 ('krasnogorsk','Красногорск','Москва'),('krasnogorsk','Москва','Красногорск'),
                 ('yasinovataya','Ясиноватая','Донецк'),('yasinovataya','Донецк','Ясиноватая')]
    pages={}
    for name,a,b in definitions:
        row=dict(origin=a,destination=b,mode='bus',basis='one_way',currency='RUB',checked_at=now(),
                 price_period='undated_published_reference_not_monthly_history',source=SOURCES[name],
                 status='source_error',items=[],statistic='mean_published_reference_fares',receipts=[])
        try:
            if name not in pages:
                start=len(session.receipts)
                response=session.get(SOURCES[name]);response.raise_for_status()
                pages[name]=(response.text,session.receipts[start:])
            html,receipts=pages[name];row['receipts']=receipts
            if name in ('donetsk','moscow'):
                items=carrier_quotes(html,a,b)
            elif name=='krasnogorsk':
                items=bank_tariffs(html)
                for item in items:
                    suburb='ЖК Изумрудные холмы' if item['number']=='1542' else 'микрорайон Чернево'
                    item.update(departure_stop=suburb if a=='Красногорск' else 'метро Тушинская',
                                arrival_stop='метро Тушинская' if a=='Красногорск' else suburb,
                                direction_method='published_route_wide_tariff')
                row['statistic']='published_official_bank_card_route_tariffs'
            else:
                items=community_tariff(html)
                for item in items:
                    item.update(departure_stop='Ясиноватая (АС)' if a=='Ясиноватая' else 'Донецк (ДС Центр)',
                                arrival_stop='Донецк (ДС Центр)' if a=='Ясиноватая' else 'Ясиноватая (АС)',
                                direction_method='published_route_wide_tariff')
                row['statistic']='secondary_community_tariff_reference'
                row['confidence']='secondary_reserve_not_operator_confirmed'
            row.update(status='ok',items=items)
        except Exception as error:
            row['error_type']=type(error).__name__
        routes.append(row)
        print(a,'->',b,row['status'],len(row['items']),flush=True)
    return dict(schema_version=2,complete=True,started_at=started,updated_at=now(),
                requested_routes=[[a,b,'bus'] for _,a,b in definitions],
                scope=sorted({n for _,a,b in definitions for n in (a,b)}),routes=routes)


def main():
    parser=argparse.ArgumentParser(__doc__)
    parser.add_argument('--raw-dir',type=Path,required=True)
    parser.add_argument('--out',type=Path,default=Path('data/surface/connecting.json'))
    args=parser.parse_args();write(args.out,collect(args.raw_dir))


if __name__=='__main__':main()
