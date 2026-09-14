#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Готовит компактный набор данных для панели задач (plan-data.js).

    python3 tools/export_dashboard.py --year 2026 --clients data/clients.csv
"""
import argparse
import calendar
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_plan import (ROOT, generate_month, load_calendar, normalize_client,  # noqa: E402
                           read_csv)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--year', type=int, default=2026)
    ap.add_argument('--clients', default=os.path.join(ROOT, 'data/clients.sample.csv'))
    ap.add_argument('--catalog', default=os.path.join(ROOT, 'data/task_catalog.csv'))
    ap.add_argument('--holidays', default=os.path.join(ROOT, 'data/holidays.csv'))
    ap.add_argument('--out', default=os.path.join(ROOT, 'dashboard/plan-data.js'))
    args = ap.parse_args()

    cal = load_calendar(args.holidays)
    clients = [normalize_client(c) for c in read_csv(args.clients)]
    catalog = read_csv(args.catalog)

    client_idx, assignee_idx, block_idx, title_idx = {}, {}, {}, {}

    def idx(store, key):
        return store.setdefault(key, len(store))

    months = {}
    for m in range(1, 13):
        rows = generate_month(args.year, m, clients, catalog, cal)
        packed = []
        for r in rows:
            packed.append([
                idx(client_idx, r['client_id']),
                idx(assignee_idx, r['assignee']),
                idx(block_idx, r['block']),
                idx(title_idx, (r['code'], r['title'].replace(r['period'], '{p}') if r['period'] else r['title'])),
                r['period'],
                int(r['internal_due'][8:10]),
                int(r['legal_due'][8:10]) if r['legal_due'] else 0,
                r['crit'],
                r['effort_min'],
                r['checklist'],
            ])
        ndays = calendar.monthrange(args.year, m)[1]
        nonworking = [d for d in range(1, ndays + 1)
                      if not cal.is_workday(dt.date(args.year, m, d))]
        months[f'{args.year}-{m:02d}'] = {'t': packed, 'days': ndays, 'off': nonworking}

    profile = {}
    for c in clients:
        chips = []
        chips.append(c['sno'])
        if c['nds']:
            chips.append('НДС')
        chips.append(f"{c['employees']} сотр." if c['employees'] else 'без сотрудников')
        for flag, label in (('kassa', 'касса'), ('ved', 'ВЭД'), ('prop', 'имущество'),
                            ('alco', 'алкоголь'), ('mark', 'маркировка'), ('op', 'обособка')):
            if c[flag]:
                chips.append(label)
        profile[c['client_id']] = {'name': c['name'], 'opf': c['opf'], 'chips': chips,
                                   'tariff': c.get('tariff', '')}

    data = {
        'year': args.year,
        'clients': [k for k, _ in sorted(client_idx.items(), key=lambda x: x[1])],
        'assignees': [k for k, _ in sorted(assignee_idx.items(), key=lambda x: x[1])],
        'blocks': [k for k, _ in sorted(block_idx.items(), key=lambda x: x[1])],
        'titles': [list(k) for k, _ in sorted(title_idx.items(), key=lambda x: x[1])],
        'profile': profile,
        'months': months,
    }
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8') as f:
        f.write('window.PLAN_DATA = ')
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print(f'{args.out}: {sum(len(v["t"]) for v in months.values())} задач, '
          f'{os.path.getsize(args.out)/1024:.0f} КБ')


if __name__ == '__main__':
    main()
