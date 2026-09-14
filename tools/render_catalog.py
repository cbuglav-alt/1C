#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Перерисовывает таблицу каталога задач в regulations/03-task-catalog.md.

Запускать после каждого изменения data/task_catalog.csv:
    python3 tools/render_catalog.py
"""
import csv
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOC = os.path.join(ROOT, 'regulations/03-task-catalog.md')
START, END = '<!-- CATALOG:START -->', '<!-- CATALOG:END -->'

FREQ = {'M': 'ежемесячно', 'Q': 'ежеквартально', 'Y': 'ежегодно', 'E': 'по событию'}
ROLE = {'ASSIST': 'Ассистент', 'ACCT': 'Бухгалтер', 'PAYROLL': 'Зарплата',
        'CHIEF': 'Главбух', 'HEAD': 'Руководитель'}
MONTHS_SHORT = ['', 'янв', 'фев', 'мар', 'апр', 'май', 'июн',
                'июл', 'авг', 'сен', 'окт', 'ноя', 'дек']


def months_label(row):
    ms = [int(x) for x in row['months'].split(',') if x.strip()]
    if not ms:
        return 'все месяцы'
    return ', '.join(MONTHS_SHORT[m] for m in ms)


def due_label(row):
    legal = int(row['legal_due_day'] or 0)
    if legal:
        lead = int(row['lead_days'] or 0)
        suffix = f' (буфер {lead} р.д.)' if lead else ''
        return f'закон: {legal} числа{suffix}'
    return f"внутренний: {row['int_due_day']} числа"


def main():
    with open(os.path.join(ROOT, 'data/task_catalog.csv'), encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f, delimiter=';'))

    out = []
    for block in dict.fromkeys(r['block'] for r in rows):
        out += [f'### {block}', '',
                '| Код | Задача | Период | Месяцы | Срок | Кому | Применяется, если | Мин |',
                '|---|---|---|---|---|---|---|---:|']
        for r in (x for x in rows if x['block'] == block):
            cond = r['applies_if'] if r['applies_if'] != '-' else 'всем клиентам'
            title = r['title'].replace('{period}', 'период')
            out.append(f"| `{r['code']}` | {title} | {FREQ[r['freq']]} | {months_label(r)} "
                       f"| {due_label(r)} | {ROLE[r['role']]} | `{cond}` | {r['effort_min']} |")
        out.append('')

    doc = open(DOC, encoding='utf-8').read()
    new = f'{START}\n\n' + '\n'.join(out).rstrip() + f'\n\n{END}'
    doc = re.sub(re.escape(START) + r'.*?' + re.escape(END), lambda _: new, doc, flags=re.S)
    open(DOC, 'w', encoding='utf-8').write(doc)
    print(f'Обновлено: {len(rows)} задач в {DOC}')


if __name__ == '__main__':
    main()
