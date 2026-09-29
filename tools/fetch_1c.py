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
    доп. реквизит «СНО» / «Доп к СНО» и теги → sno, nds, psn
    теги «Работодатель», «Только найм», заполненный «Бухгалтер ЗП» → есть сотрудники
    тег «ККТ»                        → kassa
    тег «Нулевщик» или «Доп к СНО» = нулевщик → zero (без обработки первички)
    тег «Маркетплейсы»               → mp (еженедельные выгрузки отчётов и УПД)
    тег «АУСН» + «УСН доходы» / «УСН доходы-расходы» → АУСН-Д / АУСН-ДР;
    объект не указан → АУСН (первичку запрашиваем, как при доходах минус расходы)
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

    def get(self, entity: str, select: str = '') -> list[dict]:
        params = {'$format': 'json'}
        if select:
            params['$select'] = select
        url = self.url + urllib.parse.quote(entity) + '?' + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers=self.headers)
        with urllib.request.urlopen(req, timeout=300) as resp:
            return json.load(resp)['value']


def main() -> int:
    ap = argparse.ArgumentParser(description='Выгрузка клиентов из 1С:УНФ через OData')
    ap.add_argument('--tag', default='Аутсорсинг', help='брать только контрагентов с этим тегом')
    ap.add_argument('--out', default=os.path.join(ROOT, 'data/clients.csv'))
    ap.add_argument('--advance-day', default='20', help='день аванса, если в карточке не указан')
    ap.add_argument('--salary-day', default='5', help='день зарплаты, если в карточке не указан')
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
        extra[d['Ref_Key']][props.get(d['Свойство_Key'], '')] = str(value or '').strip()

    clients = api.get('Catalog_Контрагенты', 'Ref_Key,Code,Description,ИНН,ВидКонтрагента,'
                                             'Ответственный_Key,DeletionMark,IsFolder')
    want = args.tag.lower()
    out, notes = [], []
    for c in clients:
        ref = c['Ref_Key']
        if c['IsFolder'] or c['DeletionMark'] or want not in tags[ref]:
            continue
        t, x = tags[ref], extra[ref]
        rec = {f: '' for f in FIELDS}
        rec['client_id'] = c['Code']
        rec['name'] = c['Description']
        rec['inn'] = c['ИНН']
        # все юрлица (ООО, НКО, АО) сдают отчётность организации — генератор различает ИП и ООО
        rec['opf'] = 'ИП' if c['ВидКонтрагента'] == 'ИндивидуальныйПредприниматель' or 'ип' in t else 'ООО'

        sno_raw = x.get('СНО', '')
        if 'аусн' in t:
            # тег АУСН главнее реквизита: реквизит «СНО» часто остаётся от прошлого режима
            obj = '-ДР' if 'усн доходы-расходы' in t else '-Д' if 'усн доходы' in t else ''
            if not obj:
                notes.append(f'{c["Description"]}: АУСН без объекта (доходы / доходы-расходы) — '
                             'первичка запрашивается')
            sno_raw = 'АУСН' + obj
        elif not sno_raw or sno_raw.lower() == 'найм':
            for tag, value in (('аусн', 'АУСН'), ('есхн', 'ЕСХН'), ('осно', 'ОСНО'),
                               ('усн доходы-расходы', 'УСН-ДР'), ('усн доходы', 'УСН-Д')):
                if tag in t:
                    sno_raw = value
                    break
        rec['sno'] = sno_raw if sno_raw.startswith('АУСН') else norm_sno(sno_raw) if sno_raw else ''
        add = x.get('Доп к СНО', '')
        rec['nds'] = '1' if 'ндс' in t or 'ндс' in add.lower() or rec['sno'] == 'ОСНО' else '0'
        rec['psn'] = '1' if is_patent(add) or 'патент' in t else '0'
        if rec['sno'] == 'ПСН' and rec['opf'] == 'ИП':
            rec['psn'] = '0'

        has_staff = bool({'работодатель', 'только найм'} & t) or bool(x.get('Бухгалтер ЗП'))
        rec['employees'] = '1' if has_staff else '0'
        rec['kassa'] = '1' if 'ккт' in t else '0'
        rec['zero'] = '1' if 'нулевщик' in t or 'нулевщик' in add.lower() else '0'
        rec['mp'] = '1' if 'маркетплейсы' in t else '0'
        for f in ('gph', 'ved', 'prop', 'alco', 'mark', 'op'):
            rec[f] = '0'

        rec['accountant'] = x.get('Главный бухгалтер', '')
        rec['assistant'] = x.get('Бухгалтер Первичник', '')
        rec['payroll'] = x.get('Бухгалтер ЗП', '')
        rec['manager'] = staff.get(c['Ответственный_Key'], '')
        if not rec['accountant']:
            rec['accountant'] = rec['manager']
            notes.append(f'{rec["name"]}: «Главный бухгалтер» не заполнен, задачи ушли ответственному')
        if has_staff:
            rec['advance_day'], rec['salary_day'] = args.advance_day, args.salary_day
        if not rec['sno']:
            notes.append(f'{rec["name"]}: не определена СНО (нет реквизита и тега)')
        rec['status'] = 'active'
        out.append(rec)

    out.sort(key=lambda r: r['name'])
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, FIELDS, delimiter=';')
        w.writeheader()
        w.writerows(out)

    print(f'Клиентов с тегом «{args.tag}»: {len(out)} → {args.out}')
    for n in notes:
        print('  ! ' + n)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
