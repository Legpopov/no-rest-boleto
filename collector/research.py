"""Extract published ATOR route monitoring, not monthly transaction averages.

python -m collector.research manifest.json --out data/research
Raw HTML must be downloaded and inspected separately before this step.
"""
import argparse
from datetime import date
import hashlib
from html import unescape
from html.parser import HTMLParser
import json
from pathlib import Path
import re

MONTHS = dict(zip(('января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'), range(1,13)))
ORIGINS = {'Москвы':'Москва','Санкт-Петербурга':'Санкт-Петербург','Екатеринбурга':'Екатеринбург',
           'Казани':'Казань','Новосибирска':'Новосибирск','Сочи':'Сочи',
           'Минеральных Вод':'Минеральные Воды','Калининграда':'Калининград'}
DESTINATIONS = {'Москву':'Москва','Санкт-Петербург':'Санкт-Петербург','Екатеринбург':'Екатеринбург',
                'Казань':'Казань','Новосибирск':'Новосибирск','Сочи':'Сочи',
                'Минеральные Воды':'Минеральные Воды','Калининград':'Калининград'}


def clean(s):
    return re.sub(r'\s+', ' ', unescape(s).replace('\u2060','').replace('\u200b','')).strip()


class Tables(HTMLParser):
    def __init__(self):
        super().__init__()
        self.text, self.tables, self.table, self.row, self.cell = [], [], None, None, None

    def handle_starttag(self, tag, attrs):
        if tag == 'table': self.table = {'context': clean(' '.join(self.text[-100:])), 'rows': []}
        elif tag == 'tr' and self.table is not None: self.row = []
        elif tag in ('td','th') and self.row is not None: self.cell = []
        elif tag in ('p','br','h2','h3'): self.text.append(' ')

    def handle_data(self, data):
        self.text.append(data)
        if self.cell is not None: self.cell.append(data)

    def handle_endtag(self, tag):
        if tag in ('td','th') and self.cell is not None:
            self.row.append(clean(' '.join(self.cell)))
            self.cell = None
        elif tag == 'tr' and self.row is not None:
            self.table['rows'].append(self.row)
            self.row = None
        elif tag == 'table' and self.table is not None:
            self.tables.append(self.table)
            self.table = None


def route_quotes(html, source):
    parser = Tables()
    parser.feed(html)
    records, rejected = [], []
    for table in parser.tables:
        # Bind origin and travel date to the nearest actual route-table caption.
        captions = list(re.finditer(r'Минимальная стоимость билетов из (.*?) в крупнейшие аэропорты России на (\d+) (\w+) (20\d{2}) года',table['context']))
        rows = table['rows']
        if not rows or not captions: continue
        header = rows[0]
        if 'Без багажа' not in header or 'С багажом' not in header: continue
        caption = captions[-1]
        origin = ORIGINS.get(caption[1])
        if origin is None: raise ValueError('Unmapped origin: '+caption[1])
        day, month, year = int(caption[2]), MONTHS[caption[3]], int(caption[4])
        for row in rows[1:]:
            if len(row) != len(header):
                rejected.append({'source':source['id'],'reason':'column_count','row':row})
                continue
            dest_text = re.sub(r'^в\s+', '', row[0])
            override = re.search(r'\((\d+) (\w+)\)',dest_text)
            dest_text = re.sub(r'\s*\(.*?\)', '', dest_text)
            dest = DESTINATIONS.get(dest_text)
            if dest is None: raise ValueError('Unmapped destination: '+dest_text)
            travel_date = date(year, MONTHS[override[2]] if override else month, int(override[1]) if override else day).isoformat()
            for label, baggage in (('Без багажа',False),('С багажом',True)):
                value = row[header.index(label)]
                match = re.match(r'([\d\s]+)\s*руб\.',value)
                if not match: raise ValueError('Unrecognized RUB price: '+value)
                price = int(re.sub(r'\s','',match[1]))
                if not 100 <= price <= 1000000: raise ValueError('Price out of validation bounds')
                record = {'origin':origin,'destination':dest,'travel_date':travel_date,
                          'travel_month':travel_date[:7],'mode':'air','basis':'one_way',
                          'baggage_included':baggage,'currency':'RUB','price':price,
                          'statistic':'minimum_fare_on_reference_date','source':source['url'],
                          'source_id':source['id'],'source_sha256':source['sha256'],
                          'retrieved_at':source['retrieved_at'],
                          'carrier':value[match.end():].strip()}
                record['id'] = hashlib.sha256(json.dumps({k:v for k,v in record.items() if k!='retrieved_at'},ensure_ascii=False,sort_keys=True).encode()).hexdigest()
                records.append(record)
    return records, rejected


def pooled_indices(html, source):
    """Publisher's sample-level YoY figures, never an individual-route index."""
    parser=Tables()
    parser.feed(html)
    out=[]
    for table in parser.tables:
        rows=table['rows']
        if not rows: continue
        header=rows[0]
        old=[i for i,h in enumerate(header) if '2025' in h]
        new=[i for i,h in enumerate(header) if '2026' in h]
        if len(old)!=1 or len(new)!=1: continue
        month=5 if source['id']=='ator-may-2026' else 6 if source['id']=='ator-june-2026' else None
        if month is None: continue
        for row in rows[1:]:
            if len(row)!=len(header): raise ValueError('Pooled table column mismatch')
            def price(cell):
                m=re.match(r'^[\d\s]+',cell)
                if not m: raise ValueError('Pooled price missing')
                value=int(re.sub(r'\s','',m[0]))
                if value<=0: raise ValueError('Invalid pooled price')
                return value
            previous,current=price(row[old[0]]),price(row[new[0]])
            out.append({'month':month,'previous_year':2025,'year':2026,
                        'category':row[0], 'measure':row[1] if old[0]==2 else row[0],
                        'previous_price':previous,'current_price':current,
                        'rate':current/previous-1,'scope':'publisher_route_sample',
                        'statistic':'aggregate_of_reference_date_minimum_fares',
                        'source':source['url'],'source_sha256':source['sha256']})
    return out


def reference_changes(rows, before='2026-05', after='2026-06'):
    key=lambda r:(r['origin'],r['destination'],r['baggage_included'],r['basis'],r['currency'])
    def period(month):
        result={}
        for r in rows:
            if r['travel_month']==month:
                if key(r) in result: raise ValueError('Ambiguous monthly reference quote')
                result[key(r)]=r
        return result
    old,new=period(before),period(after)
    return [{'origin':k[0],'destination':k[1],'baggage_included':k[2],
             'before_date':old[k]['travel_date'],'after_date':new[k]['travel_date'],
             'before_price':old[k]['price'],'after_price':new[k]['price'],
             'rate':new[k]['price']/old[k]['price']-1,
             'method':'reference_date_change_not_monthly_average',
             'before_id':old[k]['id'],'after_id':new[k]['id']}
            for k in sorted(old.keys() & new.keys())]


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('manifest',type=Path)
    ap.add_argument('--out',type=Path,required=True)
    args=ap.parse_args()
    all_rows, rejected, sources, indices = {}, [], [], []
    for source in json.loads(args.manifest.read_text(encoding='utf-8')):
        if source.get('status')!=200 or not source['id'].startswith('ator-'): continue
        raw=args.manifest.parent/(source['id']+'.html')
        if hashlib.sha256(raw.read_bytes()).hexdigest()!=source['sha256']:
            raise ValueError('Source digest mismatch: '+str(raw))
        rows, errors=route_quotes(raw.read_text(encoding='utf-8'),source)
        indices.extend(pooled_indices(raw.read_text(encoding='utf-8'),source))
        sources.append({**source,'extracted_quotes':len(rows)})
        rejected.extend(errors)
        all_rows.update((r['id'],r) for r in rows)
    args.out.mkdir(parents=True,exist_ok=True)
    ordered=sorted(all_rows.values(),key=lambda r:(r['travel_date'],r['origin'],r['destination'],r['baggage_included']))
    for filename, content in [('ator-quotes.json',ordered),('sources.json',sources),('rejections.json',rejected),
                              ('ator-pooled-yoy.json',indices),('ator-reference-changes.json',reference_changes(ordered))]:
        (args.out/filename).write_text(json.dumps(content,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'quotes':len(ordered),'sources':len(sources),'rejected':len(rejected),
                      'travel_months':sorted({r['travel_month'] for r in ordered})},ensure_ascii=False))


if __name__=='__main__': main()
