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
FIELDS = ['client_id', 'name', 'inn', 'opf', 'sno', 'nds', 'employees', 'gph', 'kassa',
          'ved', 'prop', 'alco', 'mark', 'op', 'accountant', 'payroll', 'chief',
          'tariff', 'advance_day', 'salary_day', 'status', 'start_date']
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


def to_int(v: str) -> str:
    m = re.search(r'\d+', (v or '').replace(' ', ''))
    return m.group(0) if m else '0'


def main() -> int:
    ap = argparse.ArgumentParser(description='Импорт заполненного справочника клиентов')
    ap.add_argument('--file', required=True)
    ap.add_argument('--out', default=os.path.join(ROOT, 'data/clients.csv'))
    args = ap.parse_args()

    rows, headers = read_rows(args.file)
    mapping = build_mapping(headers, {})
    missing = [f for f in SYNONYMS if f not in mapping]

    out, skipped, seq = [], 0, 0
    for row in rows:
        name = get(row, mapping, 'name')
        code = get(row, mapping, 'client_id')
        if not name or code.upper().startswith('ПРИМЕР') or name.upper().startswith('ПРИМЕР'):
            skipped += 1
            continue
        seq += 1
        rec = {f: '' for f in FIELDS}
        rec['client_id'] = code or f'CL-{seq:03d}'
        rec['name'] = name
        rec['inn'] = re.sub(r'\D', '', get(row, mapping, 'inn'))
        rec['opf'] = get(row, mapping, 'opf').upper()
        rec['sno'] = get(row, mapping, 'sno').upper().replace(' ', '')
        for f in BOOL_FIELDS:
            rec[f] = '1' if tri_bool(get(row, mapping, f)) else '0'
        rec['employees'] = to_int(get(row, mapping, 'employees'))
        for f in ('accountant', 'payroll', 'chief', 'tariff'):
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

    print(f'Перенесено клиентов: {len(out)}' + (f' (пропущено строк без наименования: {skipped})'
                                                if skipped else ''))
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
