#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор месячного плана задач бухгалтерского агентства.

На вход:  справочник клиентов + каталог задач + производственный календарь.
На выход: план задач на месяц (общий, по бухгалтерам, сводка, JSON, файл для импорта в CRM).

Пример:
    python3 tools/generate_plan.py --month 2026-10
    python3 tools/generate_plan.py --year 2027 --clients data/clients.csv
"""
from __future__ import annotations

import argparse
import calendar
import csv
import datetime as dt
import json
import os
import re
import sys
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

MONTHS_RU = ['', 'январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
             'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
MONTHS_RU_GEN = ['', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
                 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
QUARTERS_RU = {1: 'I квартал', 2: 'II квартал', 3: 'III квартал', 4: 'IV квартал'}
CUM_RU = {4: 'I квартал', 7: 'полугодие', 10: '9 месяцев'}

BOOL_FIELDS = ('nds', 'gph', 'kassa', 'ved', 'prop', 'alco', 'mark', 'op')
INT_FIELDS = ('employees', 'advance_day', 'salary_day')


# --------------------------------------------------------------------------- #
# Производственный календарь
# --------------------------------------------------------------------------- #
class WorkCalendar:
    def __init__(self, holidays: set[dt.date]):
        self.holidays = holidays

    def is_workday(self, d: dt.date) -> bool:
        return d.weekday() < 5 and d not in self.holidays

    def forward(self, d: dt.date) -> dt.date:
        """Перенос срока на ближайший следующий рабочий день (НК РФ ст. 6.1 п. 7)."""
        while not self.is_workday(d):
            d += dt.timedelta(days=1)
        return d

    def backward(self, d: dt.date) -> dt.date:
        while not self.is_workday(d):
            d -= dt.timedelta(days=1)
        return d

    def minus_workdays(self, d: dt.date, n: int) -> dt.date:
        d = self.backward(d)
        for _ in range(n):
            d -= dt.timedelta(days=1)
            d = self.backward(d)
        return d


def load_calendar(path: str) -> WorkCalendar:
    holidays: set[dt.date] = set()
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or line.startswith('date'):
                continue
            holidays.add(dt.date.fromisoformat(line.split(';')[0]))
    return WorkCalendar(holidays)


# --------------------------------------------------------------------------- #
# Справочники
# --------------------------------------------------------------------------- #
def read_csv(path: str) -> list[dict]:
    with open(path, encoding='utf-8-sig') as f:
        rows = [r for r in csv.DictReader(f, delimiter=';')
                if r.get('code') or r.get('client_id')]
    return rows


def normalize_client(c: dict) -> dict:
    c = dict(c)
    for k in BOOL_FIELDS:
        c[k] = str(c.get(k, '0')).strip() in ('1', 'да', 'true', 'yes')
    for k in INT_FIELDS:
        raw = str(c.get(k, '0')).strip()
        c[k] = int(raw) if raw.isdigit() else 0
    c['sno'] = (c.get('sno') or '').strip()
    c['opf'] = (c.get('opf') or '').strip()
    return c


def task_applies(expr: str, client: dict) -> bool:
    expr = (expr or '-').strip()
    if expr in ('-', ''):
        return True
    ns = {k: client[k] for k in BOOL_FIELDS + INT_FIELDS}
    ns['sno'] = client['sno']
    ns['opf'] = client['opf']
    try:
        return bool(eval(expr, {'__builtins__': {}}, ns))
    except Exception as exc:  # noqa: BLE001
        raise SystemExit(f'Ошибка в правиле применимости "{expr}": {exc}')


# --------------------------------------------------------------------------- #
# Периоды и сроки
# --------------------------------------------------------------------------- #
def period_label(period_code: str, year: int, month: int) -> str:
    if period_code == 'CM':
        return f'{MONTHS_RU[month]} {year}'
    if period_code == 'PM':
        y, m = (year - 1, 12) if month == 1 else (year, month - 1)
        return f'{MONTHS_RU[m]} {y}'
    if period_code == 'PQ':
        q = (month - 1) // 3  # квартал, предшествующий кварталу месяца исполнения
        return f'{QUARTERS_RU[4]} {year - 1}' if q == 0 else f'{QUARTERS_RU[q]} {year}'
    if period_code == 'CUM':
        return f'{CUM_RU[month]} {year}' if month in CUM_RU else f'{year - 1} год'
    if period_code == 'PY':
        return f'{year - 1} год'
    return ''


def resolve_day(token: str, client: dict, year: int, month: int) -> int | None:
    """Разбор дня внутреннего срока: число, 'last', 'ADV-2', 'SAL-1'."""
    token = (token or '').strip()
    if not token:
        return None
    if token == 'last':
        return calendar.monthrange(year, month)[1]
    m = re.fullmatch(r'(ADV|SAL)([+-]\d+)?', token)
    if m:
        base = client['advance_day'] if m.group(1) == 'ADV' else client['salary_day']
        if not base:
            return None
        day = base + int(m.group(2) or 0)
        return max(1, min(day, calendar.monthrange(year, month)[1]))
    if token.isdigit():
        return max(1, min(int(token), calendar.monthrange(year, month)[1]))
    return None


def build_task(task: dict, client: dict, year: int, month: int, cal: WorkCalendar) -> dict | None:
    months = [int(x) for x in task['months'].split(',') if x.strip()]
    if months and month not in months:
        return None
    if not task_applies(task['applies_if'], client):
        return None

    legal_day = int(task['legal_due_day'] or 0)
    lead = int(task['lead_days'] or 0)

    legal_due = None
    if legal_day:
        legal_due = cal.forward(dt.date(year, month, min(legal_day, calendar.monthrange(year, month)[1])))
        internal_due = cal.minus_workdays(legal_due, lead)
    else:
        day = resolve_day(task['int_due_day'], client, year, month)
        if day is None:
            return None
        internal_due = cal.backward(dt.date(year, month, day))

    start = cal.minus_workdays(internal_due, 3)
    first = cal.forward(dt.date(year, month, 1))
    if start < first:
        start = first

    plabel = period_label(task['period'], year, month)
    title = task['title'].replace('{period}', plabel)

    return {
        'code': task['code'],
        'client_id': client['client_id'],
        'client': client['name'],
        'title': title,
        'block': task['block'],
        'period': plabel,
        'role': task['role'],
        'assignee': assignee_for(task['role'], client),
        'start': start.isoformat(),
        'internal_due': internal_due.isoformat(),
        'legal_due': legal_due.isoformat() if legal_due else '',
        'crit': task['crit'],
        'effort_min': int(task['effort_min'] or 0),
        'checklist': task['checklist'],
        'basis': task['basis'],
        'tariff': client.get('tariff', ''),
    }


def assignee_for(role: str, client: dict) -> str:
    mapping = {
        'ASSIST': client.get('accountant', ''),
        'ACCT': client.get('accountant', ''),
        'PAYROLL': client.get('payroll') or client.get('accountant', ''),
        'CHIEF': client.get('chief', ''),
        'HEAD': client.get('chief', ''),
    }
    return mapping.get(role, client.get('accountant', ''))


# --------------------------------------------------------------------------- #
# Выгрузки
# --------------------------------------------------------------------------- #
CSV_FIELDS = ['code', 'client_id', 'client', 'title', 'block', 'period', 'role', 'assignee',
              'start', 'internal_due', 'legal_due', 'crit', 'effort_min', 'checklist', 'basis', 'tariff']


def write_csv(path: str, rows: list[dict], fields: list[str] = CSV_FIELDS) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fields, delimiter=';', extrasaction='ignore')
        w.writeheader()
        w.writerows(rows)


def write_crm_import(path: str, rows: list[dict], tag: str) -> None:
    """Плоский файл под массовое заведение задач в CRM.

    Названия колонок сопоставьте с полями импорта вашей CRM при первой загрузке.
    """
    out = []
    for r in rows:
        out.append({
            'Название задачи': f"[{r['code']}] {r['title']}",
            'Клиент': r['client'],
            'ИНН/ID клиента': r['client_id'],
            'Ответственный': r['assignee'],
            'Дата начала': r['start'],
            'Срок (внутренний)': r['internal_due'],
            'Законный срок': r['legal_due'],
            'Приоритет': {'A': 'Высокий', 'B': 'Средний', 'C': 'Низкий'}[r['crit']],
            'Категория': r['block'],
            'Период': r['period'],
            'Плановое время, мин': r['effort_min'],
            'Чек-лист': r['checklist'],
            'Основание': r['basis'],
            'Метка': tag,
        })
    write_csv(path, out, list(out[0].keys()) if out else ['Название задачи'])


def write_markdown(path: str, rows: list[dict], year: int, month: int) -> None:
    by_assignee = defaultdict(list)
    for r in rows:
        by_assignee[r['assignee'] or '— не назначен —'].append(r)

    total_h = sum(r['effort_min'] for r in rows) / 60
    lines = [f'# План задач на {MONTHS_RU[month]} {year}', '',
             f'Задач всего: **{len(rows)}** · клиентов: **{len({r["client_id"] for r in rows})}** '
             f'· плановая трудоёмкость: **{total_h:.1f} ч**', '',
             '> Сгенерировано `tools/generate_plan.py`. Внутренние сроки уже сдвинуты '
             'на рабочие дни и содержат буфер до законного срока.', '']

    lines += ['## Нагрузка по сотрудникам', '',
              '| Сотрудник | Задач | Трудоёмкость, ч | Задач категории A |',
              '|---|---:|---:|---:|']
    for name in sorted(by_assignee):
        rs = by_assignee[name]
        lines.append(f'| {name} | {len(rs)} | {sum(r["effort_min"] for r in rs)/60:.1f} '
                     f'| {sum(1 for r in rs if r["crit"] == "A")} |')
    lines.append('')

    for name in sorted(by_assignee):
        lines += [f'## {name}', '']
        by_client = defaultdict(list)
        for r in by_assignee[name]:
            by_client[r['client']].append(r)
        for client in sorted(by_client):
            lines += [f'### {client}', '',
                      '| Срок | Код | Задача | Законный срок | Кр. | Мин |',
                      '|---|---|---|---|---|---:|']
            for r in sorted(by_client[client], key=lambda x: (x['internal_due'], x['code'])):
                legal = ru_date(r['legal_due']) if r['legal_due'] else '—'
                lines.append(f'| {ru_date(r["internal_due"])} | `{r["code"]}` | {r["title"]} '
                             f'| {legal} | {r["crit"]} | {r["effort_min"]} |')
            lines.append('')

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))


def ru_date(iso: str) -> str:
    d = dt.date.fromisoformat(iso)
    return f'{d.day} {MONTHS_RU_GEN[d.month]}'


# --------------------------------------------------------------------------- #
def generate_month(year: int, month: int, clients: list[dict], catalog: list[dict],
                   cal: WorkCalendar) -> list[dict]:
    rows = []
    for client in clients:
        if (client.get('status') or 'active').strip() != 'active':
            continue
        for task in catalog:
            row = build_task(task, client, year, month, cal)
            if row:
                rows.append(row)
    rows.sort(key=lambda r: (r['assignee'], r['client'], r['internal_due'], r['code']))
    return rows


def safe_name(s: str) -> str:
    return re.sub(r'[^\w\-.]+', '_', s, flags=re.UNICODE).strip('_') or 'unassigned'


def main() -> int:
    ap = argparse.ArgumentParser(description='Генератор месячного плана задач')
    ap.add_argument('--month', help='месяц плана в формате YYYY-MM')
    ap.add_argument('--year', type=int, help='сгенерировать все 12 месяцев года')
    ap.add_argument('--clients', default=os.path.join(ROOT, 'data/clients.sample.csv'))
    ap.add_argument('--catalog', default=os.path.join(ROOT, 'data/task_catalog.csv'))
    ap.add_argument('--holidays', default=os.path.join(ROOT, 'data/holidays.csv'))
    ap.add_argument('--out', default=os.path.join(ROOT, 'out'))
    args = ap.parse_args()

    if not args.month and not args.year:
        today = dt.date.today()
        nxt = (today.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
        args.month = f'{nxt.year}-{nxt.month:02d}'
        print(f'Месяц не указан, планирую следующий: {args.month}', file=sys.stderr)

    cal = load_calendar(args.holidays)
    clients = [normalize_client(c) for c in read_csv(args.clients)]
    catalog = read_csv(args.catalog)

    periods = ([(args.year, m) for m in range(1, 13)] if args.year
               else [(int(args.month[:4]), int(args.month[5:7]))])

    for year, month in periods:
        rows = generate_month(year, month, clients, catalog, cal)
        tag = f'{year}-{month:02d}'
        write_csv(os.path.join(args.out, f'plan_{tag}.csv'), rows)
        write_markdown(os.path.join(args.out, f'plan_{tag}.md'), rows, year, month)
        write_crm_import(os.path.join(args.out, f'crm_import_{tag}.csv'), rows, tag)
        with open(os.path.join(args.out, f'plan_{tag}.json'), 'w', encoding='utf-8') as f:
            json.dump({'month': tag, 'tasks': rows}, f, ensure_ascii=False, indent=1)

        by_assignee = defaultdict(list)
        for r in rows:
            by_assignee[r['assignee']].append(r)
        for name, rs in by_assignee.items():
            write_csv(os.path.join(args.out, f'by_accountant/{tag}', f'{safe_name(name)}.csv'), rs)

        print(f'{tag}: {len(rows)} задач, {len({r["client_id"] for r in rows})} клиентов, '
              f'{sum(r["effort_min"] for r in rows)/60:.1f} ч')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
