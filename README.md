# XlsMdpParser

Репозиторий содержит два связанных инструмента для обработки Excel-файлов МДП:

| Компонент | Путь | Назначение |
|-----------|------|------------|
| **Единый Python-парсер** | `python/` | Один входной XLSX → HTML или `*_корр.xlsx` |
| **C# макрос (legacy)** | `XlsxMdpParser/` | Исходная реализация XLSX-корректора на EPPlus |

## Быстрый старт (Python)

```bash
cd python
python3 -m pip install -r requirements.txt
PYTHONPATH=src python3 run_gui.py
```

Подробности — в [`python/README.md`](python/README.md).

## Ветка unified-parser

В ветке `feature/unified-parser` добавлен единый парсер с выбором формата результата в GUI и CLI.
