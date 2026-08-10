# Единый парсер МДП

Один входной XLSX → два формата результата:

| Формат | Назначение |
|--------|------------|
| **HTML** | Автономный интерактивный файл с расчётом, копированием МДП и графиком |
| **XLSX** | Исправленный файл `*_корр.xlsx` с новым листом «Ремонтные схемы» |

Python-парсер (`python/src/mdp_converter/`) объединяет логику HTML-конвертера и XLSX-макроса из `XlsxMdpParser/Program.cs`.

## Установка

```bash
cd python
python3 -m pip install -r requirements.txt
```

## CLI

```bash
# HTML (по умолчанию)
PYTHONPATH=src python3 -m mdp_converter.cli "файл.xlsx" -o результат.html

# XLSX
PYTHONPATH=src python3 -m mdp_converter.cli "файл.xlsx" --format xlsx -o "файл_корр.xlsx"

# Папка → папка _корр или _html
PYTHONPATH=src python3 -m mdp_converter.cli "/путь/к/папке" --format xlsx -o "/путь/к/папке/_корр"
```

Или через обёртки:

```bash
python3 convert.py input.xlsx --format xlsx
python3 run_gui.py
```

## C# макрос

Исходный C#-макрос (`XlsxMdpParser/Program.cs`) сохранён для сравнения и автономной сборки .NET. Новый Python-парсер воспроизводит его ключевые правила:

- объединение контроля доп. параметров по схеме;
- «Минимальный из:» для нескольких МДП;
- скрытие пустых колонок;
- переименование старого листа в `old`.

## Тесты

```bash
cd python
PYTHONPATH=src python3 -m pytest tests/ -q
```

Для интеграционных тестов положите образцы XLSX в `python/Исходные файлы/` или укажите путь в `tests/test_converter.py`.
