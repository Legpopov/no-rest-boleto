"""Lossless dictionary/column encoding for fast offline Excel catalog import."""
from __future__ import annotations

FIELDS=('origin','destination','mode','month','base_year','price','annual_rate',
        'baggage_included','source','previous_source','reference_date','previous_date',
        'rate_method','quote_id','rail_class','firm','sample_count','price_method',
        'uncertainty','retrieved_at')
NUMERIC={'month','base_year','price','annual_rate','sample_count'}
FLAGS={'baggage_included','firm'}


def encode_catalog(catalog):
    strings=[]; positions={}; rows=[]
    for rec in catalog['records']:
        row=[]
        for key in FIELDS:
            if key in NUMERIC: value=rec.get(key,1 if key=='sample_count' else None)
            elif key in FLAGS:value=rec.get(key,False)
            else:
                value=rec.get(key,'')
                if not isinstance(value,str):raise ValueError('Expected text catalog field')
                if value not in positions:
                    positions[value]=len(strings);strings.append(value)
                value=positions[value]
            row.append(value)
        rows.append(row)
    return {**{k:v for k,v in catalog.items() if k!='records'},
            'record_encoding':'dictionary_columns_v1','record_columns':list(FIELDS),
            'strings':strings,'record_values':rows}


def decode_records(catalog):
    if catalog['record_encoding']!='dictionary_columns_v1' or catalog['record_columns']!=list(FIELDS):
        raise ValueError('Unsupported Excel record encoding')
    out=[]
    for row in catalog['record_values']:
        if len(row)!=len(FIELDS):raise ValueError('Invalid encoded row width')
        record={}
        for key,value in zip(FIELDS,row):
            if key not in NUMERIC|FLAGS:
                if isinstance(value,bool) or not isinstance(value,int) or not 0<=value<len(catalog['strings']):
                    raise ValueError('Invalid string dictionary index')
                value=catalog['strings'][value]
                if not isinstance(value,str):raise ValueError('Invalid dictionary string')
            record[key]=value
        out.append(record)
    return out
