#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Аудит заполненности Finkoper по официальным инструкциям.

Проверяет выгрузку клиентов из CRM (Настройки → Клиенты → Экспорт клиентов в XLSX)
против чек-листа regulations/10-finkoper-checklist.md и отвечает на два вопроса:

  1. Какие поля вообще не заведены в CRM — значит, соответствующая автоматизация
     Finkoper (сопоставление отчётов, привязка писем, счета, рентабельность) не работает
     ни по одному клиенту.
  2. У каких клиентов какие поля пустые — построчный список на исправление.

    python3 tools/audit_finkoper.py --file export_clients.xlsx
    python3 tools/audit_finkoper.py --file export_clients.xlsx --map mapping.json

Выход: out/audit_finkoper.md (сводка) и out/audit_finkoper.csv (список замечаний).

Зашифрованные собственные поля в экспорт не попадают — их заполненность
проверяется только в интерфейсе.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import re
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from audit_clients import norm, read_rows, tri_bool, inn_valid  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

LEVELS = ['Критично', 'Важно', 'Желательно']

# Поля чек-листа: как они могут называться в выгрузке Finkoper.
# Значение — (синонимы, уровень, что сломается, если поля нет в CRM вообще).
FIELDS = {
    'name': (
        ['name', 'наименование', 'название', 'клиент', 'организация', 'компания'],
        'Критично', 'Клиента нельзя идентифицировать'),
    'inn': (
        ['inn', 'инн'],
        'Критично',
        'Не работают: статусы из 1С-Отчётности, проверка статотчётности, '
        'автозаполнение реквизитов в документах, кнопки «Выписка / Финансы / Связи»'),
    'representative': (
        ['представитель', 'представитель клиента', 'контактное лицо', 'фио контакта', 'контакт'],
        'Важно', 'Некому адресовать запрос документов, не работает личный Telegram-бот'),
    'email': (
        ['email', 'e-mail', 'почта', 'электронная почта', 'адрес электронной почты'],
        'Критично', 'Письма не привязываются к клиенту и не попадают в его чат'),
    'phone': (
        ['phone', 'телефон', 'номер телефона', 'whatsapp'],
        'Важно', 'Чат WhatsApp не привязывается к клиенту'),
    'telegram': (
        ['telegram', 'телеграм', 'телеграмм'],
        'Важно', 'Нет канала Telegram: документы и вопросы идут мимо CRM'),
    'accountant': (
        ['accountant', 'ответственный', 'ответственные', 'бухгалтер', 'ведущий бухгалтер',
         'сотрудники', 'закреплённые сотрудники'],
        'Критично', 'Клиент не виден сотруднику, автозадачи не адресуются'),
    'payroll': (
        ['payroll', 'зарплатный бухгалтер', 'бухгалтер по зарплате', 'расчётчик'],
        'Важно', 'Зарплатные задачи уходят ведущему бухгалтеру'),
    'chief': (
        ['chief', 'главный бухгалтер', 'главбух', 'куратор'],
        'Важно', 'Не формируются задачи контроля качества'),
    'sno': (
        ['sno', 'сно', 'система налогообложения', 'налоговый режим', 'режим налогообложения'],
        'Критично', 'Налоговый блок задач не формируется'),
    'nds': (
        ['nds', 'ндс', 'плательщик ндс'],
        'Критично', 'Блок НДС не формируется'),
    'employees': (
        ['employees', 'сотрудники клиента', 'количество сотрудников', 'численность', 'штат'],
        'Критично', 'Зарплатный и кадровый блок не формируется'),
    'kassa': (['kassa', 'касса', 'ккт', 'онлайн-касса'],
              'Важно', 'Нет задач по кассовой дисциплине и сверке с ОФД'),
    'ved': (['ved', 'вэд', 'внешнеэкономическая деятельность'],
            'Важно', 'Нет задач по валютному контролю и курсовым разницам'),
    'prop': (['prop', 'имущество', 'основные средства'],
             'Важно', 'Нет имущественных налогов'),
    'mark': (['mark', 'маркировка', 'прослеживаемость', 'честный знак'],
             'Важно', 'Нет отчёта по прослеживаемым товарам'),
    'alco': (['alco', 'алкоголь', 'егаис'],
             'Желательно', 'Нет деклараций по обороту алкоголя'),
    'op': (['op', 'обособленные подразделения', 'обособка'],
           'Важно', 'НДФЛ не разносится по ОКТМО'),
    'advance_day': (['advance_day', 'день аванса', 'аванс', 'выплата аванса'],
                    'Важно', 'Не создаётся задача на расчёт аванса'),
    'salary_day': (['salary_day', 'день зарплаты', 'зарплата', 'выплата заработной платы'],
                   'Важно', 'Не создаётся задача на расчёт зарплаты'),
    'shift': (['shift', 'перенос', 'перенос даты', 'сдвиг', 'буфер'],
              'Важно', 'Задача появляется в календаре в день законного срока, '
                       'буфера на протокол из ФНС нет'),
    'patent': (['patent', 'патент', 'патенты'],
               'Важно', 'У ИП на ПСН нет автозадач на оплату, продление и уменьшение патента'),
    'contract_no': (['contract_no', 'номер договора', 'договор'],
                    'Важно', 'Не формируется счёт за месяц и не подставляются переменные в документы'),
    'contract_date': (['contract_date', 'дата договора'],
                      'Важно', 'То же'),
    'tariff': (['tariff', 'тариф', 'тариф по договору'],
               'Важно', 'Не считается отклонение факта от договора'),
    'contract_sum': (['contract_sum', 'сумма по договору', 'стоимость обслуживания', 'абонплата'],
                     'Важно', 'Рентабельность клиента не считается'),
    'tags': (['tags', 'теги', 'метки'],
             'Желательно', 'Нет быстрой выборки клиентов по признаку'),
    'status': (['status', 'статус', 'состояние'],
               'Важно', 'По расторгнутым клиентам продолжают создаваться задачи'),
    'start_date': (['start_date', 'дата начала обслуживания', 'начало обслуживания', 'дата старта'],
                   'Желательно', 'Не отделить наши периоды от периодов предыдущего бухгалтера'),
}

# Настройки уровня компании — по выгрузке не проверяются, подтверждаются вручную.
COMPANY_CHECKS = [
    ('Роли сотрудников созданы и проставлены каждому сотруднику', 'Критично',
     'Настройки → Настройки компании → Роли сотрудников; Настройки → Сотрудники'),
    ('Собственные поля заведены по списку data/finkoper_custom_fields.csv', 'Критично',
     'Настройки → Настройки компании → Собственные поля в карточке клиента'),
    ('У признаков профиля включён флаг «Обязательное»', 'Важно',
     'там же, при создании поля'),
    ('Шаблоны задач месячного контура созданы', 'Важно',
     'Настройки → Настройки компании → Шаблоны задач'),
    ('Тарифы созданы и назначены клиентам', 'Важно', 'Настройки → Тарифы'),
    ('Получатель оплаты (своя компания) заведён', 'Важно', 'Настройки, раздел счетов'),
    ('База знаний по клиентам создана (одна на бухкомпанию)', 'Важно', 'База знаний'),
    ('ПИН-коды сотрудникам выданы, поля с доступами зашифрованы', 'Важно',
     'Настройки → Сотрудники; Собственные поля'),
    ('Оклад каждого сотрудника заполнен', 'Важно',
     'Настройки → Сотрудники (иначе себестоимость часа не считается)'),
    ('Интеграция 1С-Отчётность / Контур / СБИС подключена', 'Важно',
     'Настройки → Интеграции (иначе статусы отчётов вносятся руками)'),
    ('Telegram-уведомления сотрудникам включены', 'Важно',
     'Настройки → Общие настройки → Уведомления о текущих задачах'),
]

ARCHIVED = {'archived', 'архив', 'архивный', 'закрыт', 'расторгнут', 'неактивен', 'inactive'}


# --------------------------------------------------------------------------- #
def build_mapping(headers: list[str], override: dict) -> dict:
    mapping, used = {}, set()
    lookup = {norm(h): h for h in headers}
    for field, (variants, _, _) in FIELDS.items():
        if field in override:
            mapping[field] = override[field]
            used.add(override[field])
            continue
        for v in variants:
            h = lookup.get(norm(v))
            if h and h not in used:
                mapping[field] = h
                used.add(h)
                break
    return mapping


def get(row: dict, mapping: dict, field: str) -> str:
    col = mapping.get(field)
    return (row.get(col) or '').strip() if col else ''


def has(row: dict, mapping: dict, field: str) -> bool:
    """Поле есть в выгрузке и заполнено непустым значением."""
    return bool(get(row, mapping, field))


# --------------------------------------------------------------------------- #
def audit(rows: list[dict], mapping: dict) -> tuple[list[dict], int]:
    findings = []
    active = 0

    for i, row in enumerate(rows, start=2):
        name = get(row, mapping, 'name') or f'(без наименования, строка {i})'
        if norm(get(row, mapping, 'status')) in ARCHIVED:
            continue
        active += 1

        sno = get(row, mapping, 'sno').upper().replace(' ', '')
        emp_raw = re.sub(r'\D', '', get(row, mapping, 'employees'))
        emp = int(emp_raw) if emp_raw else None

        # --- поля, которые проверяются «заполнено / не заполнено» ---
        simple = ['name', 'representative', 'email', 'accountant', 'chief', 'sno',
                  'contract_no', 'contract_date', 'tariff', 'contract_sum',
                  'shift', 'tags', 'status', 'start_date']
        for f in simple:
            if f in mapping and not has(row, mapping, f):
                findings.append({'Строка': i, 'Клиент': name, 'Уровень': FIELDS[f][1],
                                 'Поле': f, 'Замечание': 'Не заполнено',
                                 'Последствие': FIELDS[f][2]})

        # --- ИНН: наличие и контрольная сумма ---
        if 'inn' in mapping:
            inn = re.sub(r'\D', '', get(row, mapping, 'inn'))
            if not inn:
                findings.append({'Строка': i, 'Клиент': name, 'Уровень': 'Критично',
                                 'Поле': 'inn', 'Замечание': 'ИНН не заполнен',
                                 'Последствие': FIELDS['inn'][2]})
            elif not inn_valid(inn):
                findings.append({'Строка': i, 'Клиент': name, 'Уровень': 'Критично',
                                 'Поле': 'inn',
                                 'Замечание': f'ИНН «{inn}» не проходит контрольную сумму',
                                 'Последствие': 'Сопоставление клиента в 1С-Отчётности '
                                                'и статотчётности не сработает'})

        # --- признаки профиля: пусто ≠ «нет» ---
        for f in ('nds', 'kassa', 'ved', 'prop', 'mark', 'alco', 'op'):
            if f in mapping and tri_bool(get(row, mapping, f)) is None:
                findings.append({'Строка': i, 'Клиент': name, 'Уровень': FIELDS[f][1],
                                 'Поле': f,
                                 'Замечание': 'Признак не заполнен (пусто не равно «нет»)',
                                 'Последствие': FIELDS[f][2]})

        # --- численность и зависящие от неё поля ---
        if 'employees' in mapping and emp is None:
            findings.append({'Строка': i, 'Клиент': name, 'Уровень': 'Критично',
                             'Поле': 'employees', 'Замечание': 'Численность не заполнена',
                             'Последствие': FIELDS['employees'][2]})
        elif emp:
            for f in ('advance_day', 'salary_day', 'payroll'):
                if f in mapping and not has(row, mapping, f):
                    findings.append({'Строка': i, 'Клиент': name, 'Уровень': 'Критично'
                                     if f != 'payroll' else 'Важно',
                                     'Поле': f,
                                     'Замечание': f'Есть сотрудники ({emp}), поле не заполнено',
                                     'Последствие': FIELDS[f][2]})

        # --- патент ---
        if sno == 'ПСН' and 'patent' in mapping and not has(row, mapping, 'patent'):
            findings.append({'Строка': i, 'Клиент': name, 'Уровень': 'Критично',
                             'Поле': 'patent',
                             'Замечание': 'Режим ПСН, но раздел «Патенты» не заполнен',
                             'Последствие': FIELDS['patent'][2]})

        # --- ни одного канала связи ---
        channels = [f for f in ('email', 'telegram', 'phone') if f in mapping]
        if channels and not any(has(row, mapping, f) for f in channels):
            findings.append({'Строка': i, 'Клиент': name, 'Уровень': 'Критично',
                             'Поле': 'каналы',
                             'Замечание': 'Не заполнен ни один канал связи '
                                          '(почта, Telegram, телефон)',
                             'Последствие': 'Запрос документов уходит мимо CRM, '
                                            'история переписки не сохраняется'})

    return findings, active


# --------------------------------------------------------------------------- #
def main() -> int:
    ap = argparse.ArgumentParser(
        description='Аудит заполненности Finkoper по инструкциям сервиса')
    ap.add_argument('--file', required=True, help='Выгрузка клиентов: XLSX или CSV')
    ap.add_argument('--map', help='JSON с явным сопоставлением: {"email": "Почта клиента"}')
    ap.add_argument('--out', default=os.path.join(ROOT, 'out'))
    args = ap.parse_args()

    rows, headers = read_rows(args.file)
    override = json.load(open(args.map, encoding='utf-8')) if args.map else {}
    mapping = build_mapping(headers, override)

    absent = [f for f in FIELDS if f not in mapping]
    findings, active = audit(rows, mapping)

    levels = Counter(f['Уровень'] for f in findings)
    by_field = Counter(f['Поле'] for f in findings)
    clients_with = len({f['Клиент'] for f in findings})
    critical_clients = len({f['Клиент'] for f in findings if f['Уровень'] == 'Критично'})

    # Готовность: доля клиентов без критичных замечаний
    ready = active - critical_clients
    pct = round(ready / active * 100) if active else 0

    print(f'Файл: {os.path.basename(args.file)}')
    print(f'Клиентов в выгрузке: {len(rows)}, активных: {active}')
    print(f'Распознано полей чек-листа: {len(mapping)} из {len(FIELDS)}')
    print()
    if absent:
        print('НЕТ В CRM (поля не заведены — автоматизация не работает ни по одному клиенту):')
        for f in absent:
            print(f'  [{FIELDS[f][1]:10}] {f:14} — {FIELDS[f][2]}')
        print()
    print(f'Замечаний: {len(findings)} у {clients_with} клиентов из {active} активных')
    for lvl in LEVELS:
        if levels[lvl]:
            print(f'  {lvl}: {levels[lvl]}')
    print(f'\nГотовы к генерации задач без оговорок: {ready} из {active} ({pct}%)')
    if by_field:
        print('\nЧаще всего не заполнено:')
        for f, n in by_field.most_common(10):
            print(f'  {f:14} {n}')

    os.makedirs(args.out, exist_ok=True)
    csv_path = os.path.join(args.out, 'audit_finkoper.csv')
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, ['Строка', 'Клиент', 'Уровень', 'Поле', 'Замечание', 'Последствие'],
                           delimiter=';')
        w.writeheader()
        w.writerows(sorted(findings, key=lambda x: (LEVELS.index(x['Уровень']), x['Клиент'])))

    md = ['# Аудит заполненности Finkoper', '',
          f'Файл: `{os.path.basename(args.file)}` · клиентов: {len(rows)} '
          f'(активных: {active}) · дата проверки: {dt.date.today().isoformat()}', '',
          f'**Готовы к генерации задач без оговорок: {ready} из {active} ({pct}%)**', '',
          f'Замечаний: **{len(findings)}** у **{clients_with}** клиентов. '
          + ' · '.join(f'{lvl}: {levels[lvl]}' for lvl in LEVELS if levels[lvl]), '']

    if absent:
        md += ['## Полей нет в CRM', '',
               'Эти поля не найдены в выгрузке. Значит, они не заведены в Finkoper — '
               'и связанная с ними автоматизация не работает ни по одному клиенту.', '',
               '| Поле | Уровень | Что не работает |', '|---|---|---|']
        md += [f'| `{f}` | {FIELDS[f][1]} | {FIELDS[f][2]} |' for f in absent]
        md += ['', 'Как завести: Настройки → Настройки компании → Собственные поля '
                   'в карточке клиента. Готовый список полей с типами — '
                   '`data/finkoper_custom_fields.csv`.', '']

    if by_field:
        md += ['## Сводка по полям', '', '| Поле | Замечаний |', '|---|---:|']
        md += [f'| `{f}` | {n} |' for f, n in by_field.most_common()]
        md += ['']

    md += ['## Настройки уровня компании', '',
           'По выгрузке не проверяются — подтвердите вручную:', '',
           '| Проверка | Уровень | Где |', '|---|---|---|']
    md += [f'| {t} | {lvl} | {where} |' for t, lvl, where in COMPANY_CHECKS]

    md += ['', '## Замечания по клиентам', '',
           '| Клиент | Уровень | Поле | Замечание | Последствие |', '|---|---|---|---|---|']
    for f in sorted(findings, key=lambda x: (LEVELS.index(x['Уровень']), x['Клиент'])):
        md.append(f"| {f['Клиент']} | {f['Уровень']} | `{f['Поле']}` | {f['Замечание']} "
                  f"| {f['Последствие']} |")

    md += ['', '---', '',
           'Чек-лист и ссылки на инструкции: `regulations/10-finkoper-checklist.md`. '
           'Зашифрованные собственные поля в экспорт не выгружаются — проверьте их в интерфейсе.']

    md_path = os.path.join(args.out, 'audit_finkoper.md')
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(md) + '\n')

    print(f'\n{csv_path}\n{md_path}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
