import os
import re

import pandas as pd
from docx import Document
from datetime import datetime


# Ключевые слова → каноническое название подрядчика
CONTRACTOR_MAP = [
    (['НГСК', 'НОВАЯ ГАЗОВАЯ'],          'ООО «Новая Газовая Строительная Компания»'),
    (['ЛЕСНЫЕ'],                          'ООО «Лесные Технологии»'),
    (['ЮГРАНЕФТЕ', 'ЮГРАНЕФТЕСТРОЙ'],    'ООО «ЮграНефтеСтрой»'),
    (['НЕФТЕСПЕЦСТРОЙ', 'НСС'],          'ООО «НефтеСпецСтрой»'),
    (['ЕВРАКОР'],                         'АО «Евракор»'),
    (['ТРУБОПРОВОДСЕРВИС'],               'ООО ЭПЦ «Трубопроводсервис»'),
    (['ТЭКПРО'],                          'ООО «ТЭКПРО»'),
    (['СИБИТЕК'],                         'АО «Сибитек»'),
    (['ЭНЕРГОСТРОЙМОНТАЖ'],              'ООО «ЭнергоСтройМонтаж»'),
    (['СТРОЙФИНАНСГРУПП'],               'ООО «СтройФинансГрупп»'),
    (['НИПИ', 'НЕФТЕГАЗПРОЕКТ'],         'ООО НИПИ «Нефтегазпроект»'),
    (['РНГМ'],                            'АО «РНГМ-ГРУПП»'),
    (['ТЮМЕНЬВТОРСЫРЬЕ', 'ТВС'],         'ООО «ТюменьВторСырье»'),
    (['ТЮМЕНЬГЕОКОМ', 'ТГК'],            'ООО «ТюменьГеоКом»'),
    (['УРАЛГЕОГРУПП', 'УГГ'],            'ООО «УралГеоГрупп»'),
    (['ЮГРАГИДРОСТРОЙ', 'ЮГС'],          'ООО «Юграгидрострой»'),
    (['ЮГОРСКИЙ ПРОЕКТНЫЙ', 'ЮПИ'],      'ООО «Югорский проектный институт»'),
    (['РОСЭКСПО'],                        'ООО «РОСЭКСПО»'),
]


def normalize_contractor(name):
    upper = name.upper()
    for keywords, canonical in CONTRACTOR_MAP:
        if any(kw in upper for kw in keywords):
            return canonical
    return name


def normalize_date(text):
    match = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})', text)
    if match:
        day, month, year = match.groups()
        year = int(year)
        if year < 100:
            year += 2000
        try:
            return datetime(year, int(month), int(day)).strftime('%d.%m.%Y')
        except ValueError:
            pass
    return text


def extract_from_docx(file_path):
    try:
        doc = Document(file_path)
        data = {
            'Дата': '',
            'Объект': '',
            'Инженер СК': '',
            'Генподрядчик': ''
        }

        genpodr = ''
        subpodr = ''

        for table in doc.tables:
            for row in table.rows:
                cells = row.cells
                row_text_upper = ' '.join(c.text.strip() for c in cells).upper()

                # Дата
                if not data['Дата']:
                    for i in range(len(cells)):
                        if cells[i].text.strip() == 'Дата':
                            for j in range(i + 1, len(cells)):
                                next_text = cells[j].text.strip()
                                if next_text and next_text != 'Дата':
                                    data['Дата'] = normalize_date(next_text)
                                    break
                            break

                # Объект
                if not data['Объект'] and 'ОБЪЕКТ' in row_text_upper and 'СТРАНИЦА' not in row_text_upper:
                    last_obj_idx = -1
                    for i, c in enumerate(cells):
                        if c.text.strip().upper() == 'ОБЪЕКТ':
                            last_obj_idx = i
                    if last_obj_idx >= 0:
                        for j in range(last_obj_idx + 1, len(cells)):
                            cell_text = cells[j].text.strip().replace('\n', ' ')
                            if cell_text:
                                start = cell_text.find('«')
                                end = cell_text.rfind('»')
                                if start != -1 and end > start:
                                    data['Объект'] = cell_text[start + 1:end].strip()
                                else:
                                    data['Объект'] = cell_text
                                break

                # Генподрядчик
                if not genpodr and 'ГЕНПОДРЯДЧИК' in row_text_upper:
                    lt = {c.text.strip() for c in cells if 'ГЕНПОДРЯДЧИК' in c.text.upper()}
                    for cell in cells:
                        ct = cell.text.strip()
                        if ct and ct not in lt:
                            genpodr = ct
                            break

                # Субподрядчик
                if not subpodr and any(c.text.strip().upper() == 'СУБПОДРЯДЧИК' for c in cells):
                    lt = {c.text.strip() for c in cells if c.text.strip().upper() == 'СУБПОДРЯДЧИК'}
                    for cell in cells:
                        ct = cell.text.strip()
                        if ct and ct not in lt:
                            subpodr = ct
                            break

            # Инженер СК: первая ячейка предпоследней строки
            if not data['Инженер СК'] and len(table.rows) >= 2:
                val = table.rows[-2].cells[0].text.strip()
                if val:
                    data['Инженер СК'] = val

        # Выбор подрядчика: Генподрядчик, если заполнен, иначе Субподрядчик
        raw = genpodr if genpodr else subpodr
        data['Генподрядчик'] = normalize_contractor(raw)

        return data

    except Exception as e:
        return {
            'Дата': f'Ошибка: {e}',
            'Объект': '',
            'Инженер СК': '',
            'Генподрядчик': ''
        }


def _to_html(df, date_str):
    rows = []
    for i, row in enumerate(df.itertuples(index=False), 1):
        bg = ' style="background:#f0f6ff"' if i % 2 == 1 else ''
        cells = f'<td style="text-align:center">{i}</td>'
        for val in row:
            cells += f'<td>{val}</td>'
        rows.append(f'<tr{bg}>{cells}</tr>')
    return (
        '<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8">'
        f'<title>Результаты парсинга {date_str}</title>'
        '<style>'
        'body{font-family:Arial,sans-serif;font-size:11px;margin:20px}'
        'h2{font-size:13px;text-align:center;margin-bottom:10px}'
        'table{border-collapse:collapse;width:100%}'
        'th,td{border:1px solid #999;padding:4px 6px;vertical-align:top}'
        'th{background:#c6d9f1;text-align:center;font-size:11px}'
        '@media print{@page{margin:10mm}}'
        '</style></head><body>'
        f'<h2>Результаты парсинга {date_str}</h2>'
        '<table><thead><tr>'
        '<th>№</th><th>Файл</th><th>Дата</th><th>Объект</th><th>Инженер СК</th><th>Генподрядчик</th>'
        '</tr></thead><tbody>'
        + '\n'.join(rows) +
        '</tbody></table></body></html>'
    )


def process_folder(folder_path, output_path='summary', fmt='excel', log_func=print):
    all_data = []
    files = [f for f in os.listdir(folder_path)
             if f.lower().endswith('.docx') and ('отчет' in f.lower() or 'отчёт' in f.lower())]
    total = len(files)

    if total == 0:
        log_func("В папке нет файлов .docx")
        return None

    for i, file in enumerate(files, 1):
        file_path = os.path.join(folder_path, file)
        log_func(f"[{i}/{total}] Обработка: {file}")
        row = extract_from_docx(file_path)
        row['Файл'] = file
        all_data.append(row)

    df = pd.DataFrame(all_data, columns=['Файл', 'Дата', 'Объект', 'Инженер СК', 'Генподрядчик'])
    date_str = datetime.now().strftime('%d.%m.%Y')

    base = os.path.splitext(output_path)[0]
    xlsx_path = base + '.xlsx'
    html_path = base + '.html'

    df.to_excel(xlsx_path, index=False)
    log_func(f"Готово! Сохранено: {os.path.basename(xlsx_path)}")
    log_func(f"Найдено файлов: {total}")

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(_to_html(df, date_str))

    return html_path, xlsx_path

