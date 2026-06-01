# Тестовые отчёты

Сюда кладите проблемные `.docx`, чтобы Cloud Agent и CI могли их разбирать.

**Обязательно для отладки Громова:**

- `Ежедневный отчет (ЮНС) от 31.05.2026 г. (Куст 84) Громов В.Б..docx`
- при необходимости: `ЮНС_merged.docx`

```bash
cp "/путь/к/файлу.docx" test_data/
git add test_data/*.docx
git commit -m "test: add Gromov report sample"
git push origin main
```

Диагностика после push:

```bash
python3 scripts/diagnose_docx.py "test_data/Ежедневный отчет (ЮНС) от 31.05.2026 г. (Куст 84) Громов В.Б..docx"
```

Файлы только на Mac в проекте **без push** в GitHub — в облаке их **нет**.
