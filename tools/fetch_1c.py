#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Выгружает клиентов из 1С:УНФ (OData) в data/clients.csv для генератора задач.

    python3 tools/fetch_1c.py                      # клиенты с тегом «Аутсорсинг»
    python3 tools/fetch_1c.py --tag Аутсорсинг --out data/clients.csv

Подключение берётся из переменных окружения ONEC_ODATA_URL, ONEC_USER, ONEC_PASSWORD.

Откуда берутся поля профиля:
    Главное → Ответственный          → manager   (коммуникация с клиентом)
    доп. реквизит «Главный бухгалтер» → accountant (ведущий бухгалтер клиента)
    доп. реквизит «Бухгалтер Первичник» → assistant
    доп. реквизит «Бухгалтер ЗП»     → payroll
Режим берётся только из вкладки «Дополнительно», теги на него не влияют:
    «СНО»                            → sno; «Найм» → НАЙМ (только зарплатный блок, главбуха нет,
                                       его задачи уходят бухгалтеру ЗП); «АУСН» без объекта →
                                       АУСН (первичку запрашиваем, как при доходах минус расходы)
    «Доп к СНО»: НДС / ПСН / нулевщик → nds, psn, zero (ОСНО всегда с НДС)
Из тегов — только то, чего во вкладке «Дополнительно» нет:
    «Работодатель», заполненный «Бухгалтер ЗП», СНО «Найм» → есть сотрудники
    «ККТ»                            → kassa
    «Маркетплейсы»                   → mp (еженедельные выгрузки отчётов и УПД)
Количества сотрудников и дней выплаты зарплаты в карточке нет: при наличии
сотрудников ставится employees=1 и дни выплат по умолчанию (--advance-day/--salary-day).
"""
from __future__ import annotations

import argparse
import base64
import csv
import json
import os
import sys
import time
import urllib.parse
import urllib.request
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from import_clients import FIELDS, is_patent, norm_sno  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class OData:
    def __init__(self, url: str, user: str, password: str):
        self.url = url.rstrip('/') + '/'
        token = base64.b64encode(f'{user}:{password}'.encode()).decode()
        self.headers = {'Authorization': f'Basic {token}', 'Accept': 'application/json'}

    def _open(self, url: str):
        req = urllib.request.Request(url, headers=self.headers)
        return urllib.request.urlopen(req, timeout=300)

    def get(self, entity: str, select: str = '') -> list[dict]:
        """Вся таблица целиком. Сервис иногда отдаёт неполный ответ — сверяемся с $count."""
        params = {'$format': 'json'}
        if select:
            params['$select'] = select
        base = self.url + urllib.parse.quote(entity)
        url = base + '?' + urllib.parse.urlencode(params)
        for attempt in range(1, 4):
            with self._open(url) as resp:
                rows = json.load(resp)['value']
            with self._open(base + '/$count') as resp:
                expected = int(resp.read().decode().strip())
            if len(rows) == expected:
                return rows
            print(f'  {entity}: получено {len(rows)} из {expected}, повтор {attempt}', file=sys.stderr)
            time.sleep(2 * attempt)
        raise SystemExit(f'{entity}: 1С отдаёт неполные данные, план не собран')


def apply_overrides(path: str, clients: list[dict]) -> list[dict]:
    """Поправки, которых нет в 1С: строка client_id;field;value;comment на одно поле клиента."""
    if not os.path.exists(path):
        return []
    by_id = {c['client_id']: c for c in clients}
    applied = []
    with open(path, encoding='utf-8-sig') as f:
        for row in csv.DictReader(f, delimiter=';'):
            cid, field = (row.get('client_id') or '').strip(), (row.get('field') or '').strip()
            if not cid or cid.startswith('#'):
                continue
            if field not in FIELDS:
                raise SystemExit(f'{path}: неизвестное поле «{field}» у {cid}')
            client = by_id.get(cid)
            if client is None:
                applied.append(f'поправка для {cid} не применена: клиента нет в выгрузке')
                continue
            value = (row.get('value') or '').strip()
            if client[field] != value:
                applied.append(f'{client["name"]}: {field} {client[field] or "—"} → {value or "—"} '
                               f'(поправка{": " + row["comment"].strip() if row.get("comment") else ""})')
                client[field] = value
    return applied


def main() -> int:
    ap = argparse.ArgumentParser(description='Выгрузка клиентов из 1С:УНФ через OData')
    ap.add_argument('--tag', default='Аутсорсинг', help='брать только контрагентов с этим тегом')
    ap.add_argument('--out', default=os.path.join(ROOT, 'data/clients.csv'))
    ap.add_argument('--advance-day', default='20', help='день аванса, если в карточке не указан')
    ap.add_argument('--salary-day', default='5', help='день зарплаты, если в карточке не указан')
    ap.add_argument('--overrides', default=os.path.join(ROOT, 'data/overrides.csv'),
                    help='ручные поправки поверх 1С: client_id;field;value;comment')
    ap.add_argument('--include-deleted', action='store_true',
                    help='брать и контрагентов, помеченных на удаление')
    args = ap.parse_args()

    try:
        api = OData(os.environ['ONEC_ODATA_URL'], os.environ['ONEC_USER'], os.environ['ONEC_PASSWORD'])
    except KeyError as exc:
        raise SystemExit(f'Не задана переменная окружения {exc}')

    staff = {x['Ref_Key']: x['Description'] for x in api.get('Catalog_Сотрудники', 'Ref_Key,Description')}
    values = {x['Ref_Key']: x['Description']
              for x in api.get('Catalog_ЗначенияСвойствОбъектов', 'Ref_Key,Description')}
    props = {x['Ref_Key']: x['Description'].strip()
             for x in api.get('ChartOfCharacteristicTypes_ДополнительныеРеквизитыИСведения',
                              'Ref_Key,Description')}
    tag_names = {x['Ref_Key']: x['Description'].strip()
                 for x in api.get('Catalog_Теги', 'Ref_Key,Description')}

    tags = defaultdict(set)
    for t in api.get('Catalog_Контрагенты_Теги', 'Ref_Key,Тег_Key'):
        tags[t['Ref_Key']].add(tag_names.get(t['Тег_Key'], '').lower())

    extra = defaultdict(dict)
    for d in api.get('Catalog_Контрагенты_ДополнительныеРеквизиты'):
        value, kind = d['Значение'], d.get('Значение_Type', '')
        if 'Catalog_Сотрудники' in kind:
            value = staff.get(value, '')
        elif 'Catalog_ЗначенияСвойствОбъектов' in kind:
            value = values.get(value, '')
        # у реквизита может быть несколько строк (например, «Доп к СНО»: НДС и ПСН) — склеиваем
        name, value = props.get(d['Свойство_Key'], ''), str(value or '').strip()
        prev = extra[d['Ref_Key']].get(name)
        extra[d['Ref_Key']][name] = f'{prev}, {value}' if prev and value else prev or value

    clients = api.get('Catalog_Контрагенты', 'Ref_Key,Code,Description,ИНН,ВидКонтрагента,'
                                             'Ответственный_Key,DeletionMark,IsFolder')
    want = args.tag.lower()
    out, notes, deleted = [], [], []
    for c in clients:
        ref = c['Ref_Key']
        if c['IsFolder'] or want not in tags[ref]:
            continue
        if c['DeletionMark']:
            deleted.append(c['Description'])
            if not args.include_deleted:
                continue
        t, x = tags[ref], extra[ref]
        rec = {f: '' for f in FIELDS}
        rec['client_id'] = c['Code']
        rec['name'] = c['Description']
        rec['inn'] = c['ИНН']
        # все юрлица (ООО, НКО, АО) сдают отчётность организации — генератор различает ИП и ООО
        rec['opf'] = 'ИП' if c['ВидКонтрагента'] == 'ИндивидуальныйПредприниматель' or 'ип' in t else 'ООО'

        # режим и всё, что к нему относится, — только из полей вкладки «Дополнительно», не из тегов
        sno_raw = x.get('СНО', '')
        rec['sno'] = norm_sno(sno_raw) if sno_raw else ''
        hire_only = rec['sno'] == 'НАЙМ'
        if rec['sno'] == 'АУСН':
            notes.append(f'{c["Description"]}: в «СНО» не указан объект АУСН '
                         '(доходы / доходы-расходы) — первичка запрашивается')
        add = x.get('Доп к СНО', '').lower()
        rec['nds'] = '1' if 'ндс' in add or rec['sno'] == 'ОСНО' else '0'
        rec['psn'] = '1' if is_patent(add) and rec['sno'] != 'ПСН' else '0'
        rec['zero'] = '1' if 'нулев' in add else '0'

        has_staff = hire_only or 'работодатель' in t or bool(x.get('Бухгалтер ЗП'))
        rec['employees'] = '1' if has_staff else '0'
        rec['kassa'] = '1' if 'ккт' in t else '0'
        rec['mp'] = '1' if 'маркетплейсы' in t else '0'
        for f in ('gph', 'ved', 'prop', 'alco', 'mark', 'op'):
            rec[f] = '0'

        rec['accountant'] = x.get('Главный бухгалтер', '')
        rec['assistant'] = x.get('Бухгалтер Первичник', '')
        rec['payroll'] = x.get('Бухгалтер ЗП', '')
        rec['manager'] = staff.get(c['Ответственный_Key'], '')
        if not rec['accountant'] and hire_only and rec['payroll']:
            rec['accountant'] = rec['payroll']
        elif not rec['accountant']:
            rec['accountant'] = rec['manager']
            notes.append(f'{rec["name"]}: «Главный бухгалтер» не заполнен, задачи ушли ответственному')
        if has_staff:
            rec['advance_day'], rec['salary_day'] = args.advance_day, args.salary_day
        if not rec['sno']:
            notes.append(f'{rec["name"]}: поле «СНО» не заполнено — налоговые задачи по режиму не созданы')
        rec['status'] = 'active'
        out.append(rec)

    applied = apply_overrides(args.overrides, out)
    out.sort(key=lambda r: r['name'])
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, FIELDS, delimiter=';')
        w.writeheader()
        w.writerows(out)

    print(f'Клиентов с тегом «{args.tag}»: {len(out)} → {args.out}')
    if deleted:
        verb = 'включены в план' if args.include_deleted else 'пропущены (--include-deleted, чтобы взять)'
        print(f'  ! помечены на удаление: {len(deleted)}, {verb}')
    for n in notes:
        print('  ! ' + n)
    for n in applied:
        print('  ✎ ' + n)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
