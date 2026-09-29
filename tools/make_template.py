#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Собирает шаблон XLSX для заполнения справочника клиентов.

    python3 tools/make_template.py
"""
import os

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'templates', 'Справочник клиентов — шаблон.xlsx')

FONT = 'Arial'
HEAD_FILL = PatternFill('solid', fgColor='1F3A4D')
FILL_ME = PatternFill('solid', fgColor='FFF7D6')   # жёлтый — заполнять вам
EXAMPLE_FILL = PatternFill('solid', fgColor='EFF3F6')
THIN = Side(style='thin', color='BFC8D1')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

YN = ['да', 'нет']
SNO = ['ОСНО', 'УСН-Д', 'УСН-ДР', 'ПСН', 'АУСН', 'ЕСХН']
OPF = ['ООО', 'ИП']
TARIFF = ['Базовый', 'Стандарт', 'Полный']
STATUS = ['активен', 'приостановлен', 'архив']

# (заголовок, ширина, список значений, подсказка, обязательное)
COLUMNS = [
    ('Код клиента', 24, None, 'Ваш внутренний код. Если нет — оставьте пусто, проставим сами', False),
    ('Наименование', 30, None, 'Как в CRM', True),
    ('ИНН', 15, None, '10 цифр у ООО, 12 у ИП', True),
    ('Форма собственности', 19, OPF, 'ООО или ИП', True),
    ('Система налогообложения', 23, SNO, 'Основной режим. Патент к УСН отмечайте отдельно', True),
    ('Плательщик НДС', 16, YN, 'Пусто ≠ «нет». Заполните явно', True),
    ('Количество сотрудников', 21, None, 'Число. 0 если сотрудников нет', True),
    ('Договоры ГПХ', 14, YN, 'Есть ли выплаты физлицам по ГПХ', False),
    ('Касса', 9, YN, 'Зарегистрирована ли ККТ', False),
    ('ВЭД', 8, YN, 'Импорт, экспорт, валютные операции', False),
    ('Имущество', 13, YN, 'Транспорт, земля, недвижимость на балансе', False),
    ('Алкоголь', 12, YN, 'Оборот алкогольной продукции, ЕГАИС', False),
    ('Маркировка', 14, YN, 'Маркированные или прослеживаемые товары', False),
    ('Обособленные подразделения', 27, YN, 'Есть ли ОП на отдельном учёте', False),
    ('Ведущий бухгалтер', 21, None, 'ФИО — получает задачи по учёту и налогам', True),
    ('Зарплатный бухгалтер', 22, None, 'ФИО. Пусто — зарплата уйдёт ведущему', False),
    ('Главный бухгалтер', 20, None, 'ФИО — задачи контроля качества', False),
    ('Тариф', 12, TARIFF, 'Объём услуг по договору', False),
    ('День аванса', 13, None, 'Число месяца из трудового договора', False),
    ('День зарплаты', 15, None, 'Число месяца из трудового договора', False),
    ('Статус', 13, STATUS, 'Задачи генерируются только для «активен»', False),
    ('Дата начала обслуживания', 24, None, 'ДД.ММ.ГГГГ', False),
]

EXAMPLE = ['ПРИМЕР — удалите строку', 'ООО «Северный Ветер»', '7707083893', 'ООО', 'ОСНО', 'да', 14, 'да',
           'нет', 'да', 'да', 'нет', 'да', 'нет', 'Иванова А.С.', 'Петрова М.И.',
           'Смирнова О.В.', 'Полный', 20, 5, 'активен', '01.02.2023']

# обязательные поля: B..G (наименование … численность) и O (ведущий бухгалтер)
ROWS = 200


def sheet_clients(wb):
    ws = wb.create_sheet('Клиенты')
    ws.freeze_panes = 'B3'

    ws['A1'] = ('Заполняйте строки начиная с 4-й. Строка 3 — пример, её можно удалить. '
                'Жёлтые колонки обязательны: без них задачи по клиенту не сгенерируются. '
                'Начатая строка с пропущенным обязательным полем подсветится красным.')
    ws['A1'].font = Font(FONT, size=10, italic=True, color='55606B')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS))
    ws.row_dimensions[1].height = 28
    ws['A1'].alignment = Alignment(vertical='center', wrap_text=True)

    for i, (title, width, values, hint, required) in enumerate(COLUMNS, start=1):
        c = ws.cell(row=2, column=i, value=title)
        c.font = Font(FONT, size=10, bold=True, color='FFFFFF')
        c.fill = HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        c.border = BORDER
        c.comment = None
        ws.column_dimensions[get_column_letter(i)].width = width
        if values:
            dv = DataValidation(type='list', formula1=f'"{",".join(values)}"',
                                allow_blank=True, showDropDown=False)
            dv.error = 'Выберите значение из списка'
            dv.errorTitle = 'Недопустимое значение'
            ws.add_data_validation(dv)
            dv.add(f'{get_column_letter(i)}3:{get_column_letter(i)}{ROWS}')
        if required:
            for r in range(3, ROWS + 1):
                ws.cell(row=r, column=i).fill = FILL_ME
    ws.row_dimensions[2].height = 46

    # подсветка незаполненных обязательных ячеек — условным форматированием,
    # чтобы в файле не было формул, требующих пересчёта
    last = get_column_letter(len(COLUMNS))
    ws.conditional_formatting.add(
        f'B4:G{ROWS}',
        FormulaRule(formula=[f'AND(ISBLANK(B4),COUNTA($B4:${last}4)>0)'],
                    fill=PatternFill('solid', fgColor='FBD5D5'), stopIfTrue=False))
    ws.conditional_formatting.add(
        f'O4:O{ROWS}',
        FormulaRule(formula=[f'AND(ISBLANK(O4),COUNTA($B4:${last}4)>0)'],
                    fill=PatternFill('solid', fgColor='FBD5D5'), stopIfTrue=False))

    for i, v in enumerate(EXAMPLE, start=1):
        c = ws.cell(row=3, column=i, value=v)
        c.font = Font(FONT, size=10, italic=True, color='55606B')
        c.fill = EXAMPLE_FILL
        c.border = BORDER

    for r in range(3, ROWS + 1):
        for i in range(1, len(COLUMNS) + 1):
            cell = ws.cell(row=r, column=i)
            if r > 3:
                cell.font = Font(FONT, size=10)
            cell.border = BORDER
    return ws


def sheet_guide(wb):
    ws = wb.create_sheet('Инструкция', 0)
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 96

    rows = [
        ('Зачем этот файл', ''),
        ('', 'Из этих признаков собирается персональный список задач по каждому клиенту: '
             'состав отчётности, сроки и исполнитель. Признак не заполнен — соответствующий '
             'блок задач не появится.'),
        ('', ''),
        ('Как заполнять', ''),
        ('Лист «Клиенты»', 'Строка 3 — пример, удалите её после того, как разберётесь. '
                           'Свои данные вносите с 4-й строки.'),
        ('Жёлтые колонки', 'Обязательны. Без них задачи по клиенту не сгенерируются.'),
        ('Выпадающие списки', 'В колонках с фиксированными значениями выбирайте из списка, '
                              'не вписывайте текст руками.'),
        ('Колонка «Контроль»', 'Показывает, сколько обязательных полей осталось пустыми в строке.'),
        ('«да» / «нет»', 'Заполняйте явно. Пустая ячейка — это не «нет», а «неизвестно»: '
                         'она молча отключит целый блок задач.'),
        ('', ''),
        ('Что даёт каждый признак', ''),
        ('Система налогообложения', 'Декларации и авансы: УСН, прибыль, 3-НДФЛ, патент. '
                                    'Ключевое поле — от него зависит весь налоговый блок.'),
        ('Плательщик НДС', 'Декларация по НДС, уплата 1/3 ежемесячно, сверка книг покупок и продаж.'),
        ('Количество сотрудников', 'Зарплата, НДФЛ, взносы, РСВ, 6-НДФЛ, ЕФС-1, персведения, кадровые документы.'),
        ('Договоры ГПХ', 'НДФЛ, взносы и персведения даже при нулевой численности.'),
        ('Касса', 'Кассовая дисциплина, сверка с ОФД.'),
        ('ВЭД', 'Валютный контроль, ГТД, курсовые разницы, импортный НДС, ставка 0%.'),
        ('Имущество', 'Авансы и декларация по имущественным налогам.'),
        ('Алкоголь', 'Декларации по обороту алкогольной продукции.'),
        ('Маркировка', 'Отчёт по прослеживаемым товарам.'),
        ('Обособленные подразделения', 'Раздельная уплата НДФЛ и отчётность по ОКТМО.'),
        ('Форма собственности', 'Сроки деклараций, бухотчётность, фиксированные взносы ИП.'),
        ('Дни аванса и зарплаты', 'К ним привязываются сроки расчёта зарплаты. '
                                  'Без них срок встанет по умолчанию и может нарушить ТК РФ.'),
        ('', ''),
        ('Откуда брать данные', ''),
        ('Система налогообложения', 'Уведомление о переходе на спецрежим, налоговая декларация '
                                    'за прошлый год, личный кабинет клиента на nalog.gov.ru. '
                                    'В 1С:УНФ этого признака по контрагентам нет.'),
        ('НДС', 'Последняя декларация по НДС либо уведомление об освобождении по ст. 145 НК РФ.'),
        ('Численность', 'Последний сданный РСВ или кадровые документы.'),
        ('ИНН, наименование', 'Выгрузка контрагентов из 1С или из CRM.'),
        ('', ''),
        ('Что дальше', ''),
        ('', 'Пришлите заполненный файл — на его основе соберём план задач на месяц '
             'по каждому клиенту с внутренними и законными сроками и распределением по сотрудникам.'),
    ]
    ws['A1'] = 'Справочник клиентов — шаблон для заполнения'
    ws['A1'].font = Font(FONT, size=14, bold=True, color='1F3A4D')
    ws.merge_cells('A1:B1')
    ws.row_dimensions[1].height = 24

    r = 3
    for a, b in rows:
        if a and not b:
            ws.cell(row=r, column=1, value=a).font = Font(FONT, size=11, bold=True, color='1F3A4D')
        else:
            ws.cell(row=r, column=1, value=a).font = Font(FONT, size=10, bold=True)
            c = ws.cell(row=r, column=2, value=b)
            c.font = Font(FONT, size=10)
            c.alignment = Alignment(wrap_text=True, vertical='top')
            ws.row_dimensions[r].height = max(15, 13 * (len(b) // 95 + 1))
        r += 1
    return ws


def main():
    wb = Workbook()
    wb.remove(wb.active)
    sheet_clients(wb)
    sheet_guide(wb)
    wb.active = 0
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print(OUT)


if __name__ == '__main__':
    main()
