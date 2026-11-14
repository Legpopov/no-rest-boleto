"""Offline semester routing: transport preference, known transfers, ticket cost.

Routes are budget scenarios, not timetable itineraries. Unknown internal airline
connections are allowed only with an explicit unverified-limit warning. Never add
an unpriced edge or interpret missing transfer expenses as confirmed zero costs.
"""
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
import math

from .planning import positive, semester_months


def money(value):
    return float(Decimal(str(value)).quantize(Decimal('.01'), rounding=ROUND_HALF_UP))


def best_route(edges, origin, destination, preferences, max_transfers=2):
    if origin == destination or not origin or not destination:
        raise ValueError('Supply two distinct cities')
    if not preferences or len(set(preferences)) != len(preferences) or any(m not in ('rail','air','bus') for m in preferences):
        raise ValueError('Supply distinct supported transport preferences')
    if not isinstance(max_transfers,int) or isinstance(max_transfers,bool) or not 0 <= max_transfers <= 2:
        raise ValueError('Supported transfer limit is 0..2')
    ranks={m:i for i,m in enumerate(preferences)}
    outgoing=defaultdict(list)
    for edge in edges:
        if edge['mode'] not in ranks or not positive(edge.get('price')):
            continue
        count=edge.get('transfers')
        if count is not None and (isinstance(count,bool) or not isinstance(count,int) or count < 0):
            raise ValueError('Invalid internal transfer count')
        outgoing[edge['origin']].append(edge)
    best=None
    def search(city, legs, visited, known, unknown, used_modes, total):
        nonlocal best
        for edge in outgoing.get(city,()):
            target=edge['destination']
            if target in visited:
                continue
            if target != destination and target not in outgoing:
                continue
            # Joining separately priced tickets is itself one connection.
            count=known+(len(legs)>0)+(edge.get('transfers') or 0)
            if count > max_transfers:
                continue
            used=used_modes | {edge['mode']}
            priority=tuple(int(mode in used) for mode in reversed(preferences[1:]))
            if best is not None and priority > best[0][0]:
                continue
            path=legs+[edge]
            price=money(total+edge['price'])
            if best is not None and priority == best[0][0]:
                if count>best[0][1] or count==best[0][1] and price>best[0][2]:
                    continue
            uncertain=unknown or edge.get('transfers') is None or edge.get('transfers_unknown',False)
            if target == destination:
                score=(priority,count,price,tuple((e['origin'],e['destination'],e['mode']) for e in path))
                if best is None or score < best[0]:
                    best=(score,path,uncertain)
            elif len(path) <= max_transfers:
                search(target,path,visited|{target},count,uncertain,used,price)
    search(origin,[],{origin},0,False,set(),0)
    if best is None:
        return dict(status='insufficient_data',price=None,legs=[],reason='No complete priced path within the known-transfer limit')
    score,legs,unknown=best
    joined=len(legs)>1
    warnings=[]
    if unknown:
        warnings.append('Internal airline transfers are unknown for part of the semester; the total transfer limit is not verified')
    if joined or unknown or score[1]>0:
        warnings.append('Ticket sum only: local transfer expenses, connection times and overnight needs are not assessed')
    return dict(status='ok',price=score[2],legs=legs,
                cities=[origin]+[e['destination'] for e in legs],
                known_transfers=score[1],transfer_limit_verified=not unknown,
                max_transfers=max_transfers,ticket_sum_only=bool(warnings),warnings=warnings,
                ranking='avoid_less_preferred_modes_then_known_transfer_lower_bound_then_ticket_sum')


def rail_allowed(row, classes):
    return any(token==row.get('rail_class') or
               token==row.get('rail_class','')+':firm' and row.get('firm') is True
               for token in classes)


class SemesterRouter:
    def __init__(self,catalog):
        self.catalog=catalog
        self.groups=defaultdict(list)
        for row in catalog['records']:
            self.groups[row['origin'],row['destination'],row['mode']].append(row)

    def edges(self,months,*,baggage=False,allowance_ru=None,allowance_ab=None,rail_classes=('coupe',),target=None):
        origins={key[0] for key in self.groups}
        countries=self.catalog.get('city_countries',{})
        output=[]
        for (a,b,mode),rows in self.groups.items():
            if target is not None and b != target and b not in origins:
                continue
            allowance=allowance_ru if countries.get(a)==countries.get(b)=='RU' else allowance_ab
            if not countries.get(a) or not countries.get(b):
                allowance=None
            estimates=[];sources=set();known=0;unknown=False
            for period in months:
                year,month=map(int,period.split('-'))
                candidates=[]
                for row in rows:
                    if row['month']!=month or row['base_year']>year:
                        continue
                    current=mode=='air' and row['rate_method'].startswith('air_current_')
                    if mode=='rail' and not rail_allowed(row,rail_classes):
                        continue
                    if mode=='air':
                        if current and baggage and (not isinstance(allowance,(float,int)) or isinstance(allowance,bool) or not math.isfinite(allowance) or allowance<0):
                            continue
                        if not current and row['baggage_included']!=baggage:
                            continue
                    candidates.append((2 if current else 1,row,current))
                if not candidates:
                    break
                priority=max(v[0] for v in candidates)
                chosen=[v for v in candidates if v[0]==priority]
                if mode!='rail' and len(chosen)!=1:
                    break
                weighted=weight=0
                valid=True
                for _,row,current in chosen:
                    sample=row.get('sample_count',1)
                    rate=row['annual_rate']
                    if not positive(row['price']) or not positive(sample) or not math.isfinite(rate) or rate<=-1 or not row.get('source'):
                        valid=False;break
                    try:
                        price=(row['price']+(allowance if current and baggage else 0))*(1+rate)**(year-row['base_year'])
                    except OverflowError:
                        valid=False;break
                    if not positive(price):
                        valid=False;break
                    weighted+=price*sample;weight+=sample
                    count=row.get('transfers') if mode=='air' else 0
                    if count is None:
                        unknown=True
                    elif isinstance(count,bool) or not isinstance(count,int) or count<0:
                        valid=False;break
                    else:
                        known=max(known,count)
                    sources.add(row['source'])
                if not valid or weight<=0:
                    break
                estimates.append(weighted/weight)
            if len(estimates)==len(months) and estimates:
                output.append(dict(origin=a,destination=b,mode=mode,price=money(sum(estimates)/len(estimates)),
                                   transfers=known,transfers_unknown=unknown,sources=sorted(sources),
                                   months=list(months)))
        return output

    def plan(self,origin,destination,launch,cohort,semester,preferences,*,calendar=None,**filters):
        months=semester_months(launch,cohort,semester,**({'calendar':calendar} if calendar else {}))
        edges=self.edges(months,target=destination,**filters)
        return best_route(edges,origin,destination,preferences)
