#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Превращает заполненный шаблон (XLSX/CSV) в data/clients.csv для генератора задач.

    python3 tools/import_clients.py --file "Справочник клиентов.xlsx"
    python3 tools/import_clients.py --file export.csv --out data/clients.csv

Колонки распознаются по русским заголовкам теми же правилами, что и в аудите,
поэтому подходит и выгрузка из CRM, а не только наш шаблон.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from audit_clients import SYNONYMS, build_mapping, get, read_rows, tri_bool  # noqa: E402
from generate_plan import BOOL_FIELDS  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIELDS = ['client_id', 'name', 'inn', 'opf', 'sno', 'nds', 'psn', 'employees', 'gph', 'kassa',
          'ved', 'prop', 'alco', 'mark', 'op', 'zero', 'mp', 'accountant', 'assistant', 'payroll', 'chief', 'manager',
          'tariff', 'advance_day', 'salary_day', 'status', 'start_date']

# Написания режимов, которые встречаются в базах: «УСН д», «УСН д-р», «УСН 6%» и т.п.
SNO_RULES = [
    # объект АУСН меняет состав задач: при «доходах» первичку не запрашиваем
    (r'аусн.*(д\s*[-/]\s*р|дох.*расх|20)', 'АУСН-ДР'),
    (r'аусн.*(дох|\bд\b|8)', 'АУСН-Д'),
    (r'аусн', 'АУСН'),
    (r'есхн', 'ЕСХН'),
    (r'нпд|самозанят', 'НПД'),
    (r'усн.*(д\s*[-/]\s*р|дох.*расх|15)', 'УСН-ДР'),
    (r'усн', 'УСН-Д'),
    (r'псн|патент', 'ПСН'),
    (r'осно|осн\b|общая', 'ОСНО'),
]
OPF_RULES = [
    (r'индивидуальн|\bип\b|предпринимател', 'ИП'),
    (r'\bооо\b|обществ.*ограничен', 'ООО'),
    (r'\bао\b|акционерн', 'АО'),
]
STATUS = {'активен': 'active', 'активный': 'active', 'действующий': 'active',
          'приостановлен': 'paused', 'пауза': 'paused',
          'архив': 'archived', 'архивный': 'archived', 'расторгнут': 'archived'}


def to_date(v: str) -> str:
    v = (v or '').strip()
    if not v:
        return ''
    v = v.split(' ')[0]
    for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%d.%m.%y'):
        try:
            return dt.datetime.strptime(v, fmt).date().isoformat()
        except ValueError:
            continue
    return v


def norm_sno(v: str) -> str:
    t = (v or '').strip().lower().replace('ё', 'е')
    for pattern, value in SNO_RULES:
        if re.search(pattern, t):
            return value
    return (v or '').strip().upper().replace(' ', '')


def norm_opf(v: str, name: str) -> str:
    for source in (v, name):
        t = (source or '').strip().lower()
        for pattern, value in OPF_RULES:
            if re.search(pattern, t):
                return value
    return (v or '').strip().upper()


def is_patent(v: str) -> bool:
    """«Доп к СНО» со значением «ПСН»/«Патент» означает совмещение с патентом."""
    return bool(re.search(r'псн|патент', (v or '').lower()))


def to_int(v: str) -> str:
    m = re.search(r'\d+', (v or '').replace(' ', ''))
    return m.group(0) if m else '0'


def main() -> int:
    ap = argparse.ArgumentParser(description='Импорт заполненного справочника клиентов')
    ap.add_argument('--file', required=True)
    ap.add_argument('--out', default=os.path.join(ROOT, 'data/clients.csv'))
    ap.add_argument('--tag', help='взять только клиентов с этим тегом, например «Аутсорсинг»')
    args = ap.parse_args()

    rows, headers = read_rows(args.file)
    mapping = build_mapping(headers, {})
    missing = [f for f in SYNONYMS if f not in mapping]

    out, skipped, filtered, seq = [], 0, 0, 0
    for row in rows:
        name = get(row, mapping, 'name')
        code = get(row, mapping, 'client_id')
        if not name or code.upper().startswith('ПРИМЕР') or name.upper().startswith('ПРИМЕР'):
            skipped += 1
            continue
        if args.tag:
            tags = get(row, mapping, 'tags')
            if args.tag.lower() not in tags.lower():
                filtered += 1
                continue
        seq += 1
        rec = {f: '' for f in FIELDS}
        rec['client_id'] = code or f'CL-{seq:03d}'
        rec['name'] = name
        rec['inn'] = re.sub(r'\D', '', get(row, mapping, 'inn'))
        rec['opf'] = norm_opf(get(row, mapping, 'opf'), name)
        rec['sno'] = norm_sno(get(row, mapping, 'sno'))
        for f in BOOL_FIELDS:
            rec[f] = '1' if tri_bool(get(row, mapping, f)) else '0'
        # патент отмечается либо галочкой, либо текстом в поле «Доп к СНО»
        if is_patent(get(row, mapping, 'psn')):
            rec['psn'] = '1'
        rec['employees'] = to_int(get(row, mapping, 'employees'))
        for f in ('accountant', 'assistant', 'payroll', 'chief', 'manager', 'tariff'):
            rec[f] = get(row, mapping, f)
        rec['advance_day'] = to_int(get(row, mapping, 'advance_day'))
        rec['salary_day'] = to_int(get(row, mapping, 'salary_day'))
        rec['status'] = STATUS.get(get(row, mapping, 'status').lower(), 'active')
        rec['start_date'] = to_date(get(row, mapping, 'start_date'))
        out.append(rec)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, FIELDS, delimiter=';')
        w.writeheader()
        w.writerows(out)

    tail = []
    if filtered:
        tail.append(f'отсеяно без тега «{args.tag}»: {filtered}')
    if skipped:
        tail.append(f'пропущено строк без наименования: {skipped}')
    print(f'Перенесено клиентов: {len(out)}' + (f' ({", ".join(tail)})' if tail else ''))
    if args.tag and 'tags' not in mapping:
        print(f'ВНИМАНИЕ: колонки с тегами в файле нет, фильтр «{args.tag}» не применён.')
    if missing:
        print('Не нашлось колонок: ' + ', '.join(missing))
        print('Эти признаки записаны как «нет» — проверьте, так ли это.')
    print(args.out)
    print('\nДальше:')
    print('  python3 tools/audit_clients.py --file ' + args.out)
    print('  python3 tools/generate_plan.py --month <YYYY-MM> --clients ' + args.out)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
