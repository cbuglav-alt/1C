#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Страница «План задач» для claude.ai: задачи месяца, фильтры, скачивание Excel и импорта.

    python3 tools/build_page.py --month 2026-10

Берёт out/plan_<месяц>.csv, out/crm_import_<месяц>.csv, Excel из export_xlsx.py,
data/clients.csv и замечания out/fetch_notes.txt. Пишет out/plan_page_<месяц>.html —
в нём реальные данные клиентов, поэтому файл не коммитится.
"""
from __future__ import annotations

import argparse
import base64
import csv
import datetime as dt
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MONTHS_RU = ['', 'январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
             'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']


def read(path: str) -> list[dict]:
    with open(path, encoding='utf-8-sig') as f:
        return list(csv.DictReader(f, delimiter=';'))


def main() -> int:
    ap = argparse.ArgumentParser(description='Страница плана задач')
    ap.add_argument('--month', required=True, help='YYYY-MM')
    ap.add_argument('--clients', default=os.path.join(ROOT, 'data/clients.csv'))
    ap.add_argument('--out', default=os.path.join(ROOT, 'out'))
    ap.add_argument('--refresh-trigger', default=os.environ.get('PLAN_REFRESH_TRIGGER', ''),
                    help='id Routine, который пересобирает план (кнопка «Обновить из 1С»)')
    args = ap.parse_args()

    year, month = int(args.month[:4]), int(args.month[5:7])
    rows = read(os.path.join(args.out, f'plan_{args.month}.csv'))
    clients = {c['client_id']: c for c in read(args.clients)}

    strings, index = [], {}

    def s(v: str) -> int:
        if v not in index:
            index[v] = len(strings)
            strings.append(v)
        return index[v]

    tasks = [[s(r['client_id']), s(r['code']), s(r['title']), s(r['block']), s(r['assignee']),
              r['start'][5:], r['internal_due'][5:], r['legal_due'][5:], r['crit'],
              int(r['effort_min'] or 0)] for r in rows]
    cl = {cid: [c['name'], c['inn'], c['opf'], c['sno'], c['nds'], c['psn'], c['employees'],
                c['kassa'], c.get('zero', ''), c.get('mp', ''), c['accountant'], c['assistant'],
                c['payroll'], c['manager']] for cid, c in clients.items()}
    notes_path = os.path.join(args.out, 'fetch_notes.txt')
    notes = []
    if os.path.exists(notes_path):
        with open(notes_path, encoding='utf-8') as f:
            notes = [line.strip() for line in f if line.strip()]
    msk = dt.datetime.now(dt.timezone(dt.timedelta(hours=3)))
    data = {'s': strings, 't': tasks, 'c': cl, 'n': notes, 'y': year, 'm': month,
            'at': msk.strftime('%d.%m.%Y %H:%M'), 'trig': args.refresh_trigger}

    xlsx_name = f'Задачи_{MONTHS_RU[month]}_{year}_по_клиентам.xlsx'
    crm_name = f'crm_import_{args.month}.csv'
    with open(os.path.join(args.out, xlsx_name), 'rb') as f:
        xlsx = base64.b64encode(f.read()).decode()
    with open(os.path.join(args.out, crm_name), 'rb') as f:
        crm = base64.b64encode(f.read()).decode()

    with open(os.path.join(ROOT, 'tools/page_template.html'), encoding='utf-8') as f:
        html = f.read()
    html = (html.replace('/*DATA*/null', json.dumps(data, ensure_ascii=False, separators=(',', ':')))
                .replace('__XLSX__', xlsx).replace('__CRM__', crm)
                .replace('__XLSX_NAME__', xlsx_name).replace('__CRM_NAME__', crm_name)
                .replace('__MONTH_TITLE__', f'{MONTHS_RU[month]} {year}'))
    path = os.path.join(args.out, f'plan_page_{args.month}.html')
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'{len(rows)} задач, {len(cl)} клиентов → {path} ({len(html) // 1024} КБ)')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
