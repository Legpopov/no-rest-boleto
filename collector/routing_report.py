"""Measure priced paths for an explicit semester, not physical reachability."""
import argparse
from collections import Counter,defaultdict
import json
from pathlib import Path

from .planning import semester_months
from .routing import SemesterRouter,best_route
from .route_campaign import read,write
from .route_scope import route_scope
from .branch_scope import endpoints


def main():
    parser=argparse.ArgumentParser(__doc__)
    parser.add_argument('--launch',type=int,default=2026)
    parser.add_argument('--cohort',type=int,default=1)
    parser.add_argument('--semester',type=int,default=2)
    parser.add_argument('--out',type=Path,default=Path('data/route-coverage/routing.json'))
    args=parser.parse_args()
    catalog=read('data/planning.json');scope=route_scope(endpoints())
    months=semester_months(args.launch,args.cohort,args.semester)
    router=SemesterRouter(catalog);edges=router.edges(months)
    direct=defaultdict(list)
    for edge in edges:direct[edge['origin'],edge['destination']].append(edge)
    preferences=['rail','air','bus']
    observed={(r['origin'],r['destination']) for r in catalog['records']}
    counts=Counter();rows=[]
    for a,b in scope['directed_pairs']:
        single=best_route(direct.get((a,b),[]),a,b,preferences)
        result=best_route(edges,a,b,preferences)
        counts['requested']+=1
        if single['status']=='ok':counts['direct_complete_semester']+=1
        if result['status']=='ok':
            counts['priced_path']+=1
            if single['status']!='ok':counts['new_vs_direct_complete_semester']+=1
            if (a,b) not in observed:counts['new_vs_no_reference_at_all']+=1
            if not result['transfer_limit_verified']:counts['unverified_transfer_limit']+=1
        else:counts['no_priced_path']+=1
        rows.append(dict(origin=a,destination=b,direct_complete=single['status']=='ok',**result))
    root=dict(snapshot_at=catalog['retrieved_at'],months=months,preferences=preferences,
              baggage=False,rail_classes=['coupe'],max_known_transfers=2,
              scope='50 official origins and the reviewed 60 destinations; directed pairs',
              limitations='Budget scenarios, not timetable feasibility. Only a ticket sum; extra transfer and overnight costs are not assessed. Unknown flight connections make the limit unverified.',
              counts=dict(counts),routes=rows)
    write(args.out,root)
    lines=['# Автоматический подбор: покрытие ценовыми путями','',
           f"Снимок: {catalog['retrieved_at']}. Семестр: {', '.join(months)}.",'',
           'Параметры проверки: поезд → авиа → автобус; купе; без запроса багажа; не больше двух известных пересадок.','',
           f"- Проверено направленных пар: {counts['requested']}.",
           f"- Полный семестр у одного участка: {counts['direct_complete_semester']}.",
           f"- Ценовой путь найден, включая пересадки: {counts['priced_path']}.",
           f"- Добавлено по сравнению с полным прямым расчётом: {counts['new_vs_direct_complete_semester']}.",
           f"- Из ранее не имевших вообще ни одной цены закрыто: {counts['new_vs_no_reference_at_all']}.",
           f"- Осталось без ценового пути: {counts['no_priced_path']}.",
           f"- Из выбранных путей предел авиапересадок не проверен: {counts['unverified_transfer_limit']}.",'',
           'Это результаты при указанных фильтрах и семестре. Они не доказывают осуществимость поездки по расписанию и не включают неопределённые дополнительные расходы.','',
           '## Оставшиеся пробелы','']
    missing=defaultdict(list)
    for row in rows:
        if row['status']!='ok':missing[row['origin']].append(row['destination'])
    for city in scope['origins']:
        if missing[city]:lines.append(f"- **{city}** → {'; '.join(missing[city])}.")
    args.out.with_suffix('.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    print(json.dumps(counts,ensure_ascii=False),flush=True)


if __name__=='__main__':main()
