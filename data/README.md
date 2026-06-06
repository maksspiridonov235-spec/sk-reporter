# Данные SK-Reporter / блок инженера

| Папка | Содержимое |
|-------|------------|
| `templates/` | Болванки подрядчиков для сборки (существующий SK-Reporter) |
| `personnel/` | Справочник персонала (`spravochnik.xlsx` → `personnel.yaml`) |
| `luvr/` | Лист учёта времени (`luvr.xlsx`) |
| `projects/{id}/` | ВОР и материалы по объекту + `project.yaml`, `vor.json` |
| `tk/` | Технологические карты ОТКК (*.doc) + `manifest.yaml` |

Тяжёлые бинарники (pdf, doc, xlsx) по умолчанию **не в git** — см. `.gitignore`.
На рабочей машине кладите файлы в эти же пути после `git pull`.
