"""Детальный анализ высот строк в шаблоне и отчёте после применения."""

import sys
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

def get_row_height_info(doc_path: Path):
    """Возвращает детальную информацию о высотах строк."""
    doc = Document(str(doc_path))
    result = []
    
    for ti, table in enumerate(doc.tables):
        tbl_info = {
            'table_idx': ti,
            'total_rows': len(table.rows),
            'rows': []
        }
        
        for ri, row in enumerate(table.rows):
            tr = row._tr
            trPr = tr.find(qn('w:trPr'))
            
            height_val = None
            height_rule = None
            
            if trPr is not None:
                trH = trPr.find(qn('w:trHeight'))
                if trH is not None:
                    height_val = trH.get(qn('w:val'))
                    height_rule = trH.get(qn('w:hRule'))
            
            # Считаем количество ячеек и проверяем объединения
            tcs = tr.findall(qn('w:tc'))
            merged_cells = 0
            for tc in tcs:
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is not None:
                    vMerge = tcPr.find(qn('w:vMerge'))
                    gridSpan = tcPr.find(qn('w:gridSpan'))
                    if vMerge is not None or gridSpan is not None:
                        merged_cells += 1
            
            tbl_info['rows'].append({
                'row_idx': ri,
                'height': height_val,
                'hRule': height_rule,
                'cells': len(tcs),
                'merged_cells': merged_cells
            })
        
        result.append(tbl_info)
    
    return result


def compare_with_template(report_info: list, template_info: list):
    """Сравнивает отчёт с шаблоном и находит расхождения."""
    print("\n" + "="*80)
    print("🔍 СРАВНЕНИЕ С ШАБЛОНОМ")
    print("="*80)
    
    for ti, tbl in enumerate(report_info):
        if ti >= len(template_info):
            print(f"\n❌ Таблица {ti}: НЕТ В ШАБЛОНЕ")
            continue
        
        template_tbl = template_info[ti]
        template_rows = template_tbl['rows']
        report_rows = tbl['rows']
        
        print(f"\n📊 Таблица {ti} (в шаблоне {len(template_rows)} строк, в отчёте {len(report_rows)} строк)")
        
        problems = []
        for ri, row in enumerate(report_rows):
            if ri >= len(template_rows):
                problems.append(f"  Строка {ri}: НЕТ В ШАБЛОНЕ (h={row['height']}, rule={row['hRule']})")
                continue
            
            template_row = template_rows[ri]
            
            # Сравниваем высоту
            if row['height'] != template_row['height']:
                problems.append(f"  Строка {ri}: ВЫСОТА {row['height']} (шаблон: {template_row['height']})")
            
            # Сравниваем hRule
            if row['hRule'] != template_row['hRule']:
                problems.append(f"  Строка {ri}: hRule={row['hRule']} (шаблон: {template_row['hRule']})")
            
            # Проверяем количество ячеек
            if row['cells'] != template_row['cells']:
                problems.append(f"  Строка {ri}: ЯЧЕЕК {row['cells']} (шаблон: {template_row['cells']}) {('← ОБЪЕДИНЕНИЯ' if row['merged_cells'] > 0 else '')}")
        
        if problems:
            print("   Проблемы:")
            for p in problems[:20]:  # Показываем первые 20 проблем
                print(p)
            if len(problems) > 20:
                print(f"   ... и ещё {len(problems) - 20} проблем")
        else:
            print("   ✅ Нет проблем")


def main():
    if len(sys.argv) < 2:
        print("Использование: python3 debug_heights.py <report.docx>")
        sys.exit(1)
    
    report_path = Path(sys.argv[1])
    template_path = Path("contractor_report/болванки (шаблоны не вырезать только копировать)/Ежедневный отчет Шаблон.docx")
    
    if not report_path.exists():
        print(f"❌ Файл не найден: {report_path}")
        sys.exit(1)
    
    if not template_path.exists():
        print(f"❌ Шаблон не найден: {template_path}")
        sys.exit(1)
    
    print("="*80)
    print(f"📄 ОТЧЁТ: {report_path.name}")
    print("="*80)
    
    report_info = get_row_height_info(report_path)
    template_info = get_row_height_info(template_path)
    
    # Выводим сводку
    print(f"\n📊 В отчёте таблиц: {len(report_info)}")
    print(f"📊 В шаблоне таблиц: {len(template_info)}")
    
    for ti, tbl in enumerate(report_info):
        print(f"\n--- Таблица {ti} ---")
        print(f"Строк: {tbl['total_rows']}")
        
        # Группируем по высоте
        height_groups = {}
        for row in tbl['rows']:
            h = row['height'] or 'None'
            rule = row['hRule'] or 'None'
            key = f"{h}({rule})"
            if key not in height_groups:
                height_groups[key] = []
            height_groups[key].append(row['row_idx'])
        
        print("Высоты строк:")
        for key, rows in sorted(height_groups.items()):
            if len(rows) <= 5:
                print(f"  {key}: строки {rows}")
            else:
                print(f"  {key}: {len(rows)} строк")
    
    # Сравниваем с шаблоном
    compare_with_template(report_info, template_info)
    
    # Сохраняем лог в файл
    import json
    log_path = report_path.parent / f"{report_path.stem}_debug.json"
    with open(log_path, 'w', encoding='utf-8') as f:
        json.dump({
            'report': report_info,
            'template': template_info
        }, f, ensure_ascii=False, indent=2)
    
    print(f"\n💾 Лог сохранён: {log_path}")


if __name__ == "__main__":
    main()
