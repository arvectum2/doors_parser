# doors_parser_final — v14 Hybrid Client Ready

Проект собирает каталог дверей по целевым линейкам из `data/target_lines.csv`.

## Версия v14 — гибридная сборка

- 6 производителей взяты без изменений из `final_ready_v11` (базовая версия, уже отправленная заказчику):
  - Скамбио Порте, Лорд, Аэлита, Комфорт, Двери регионов, Термодур
- 4 проблемных производителя взяты из `final_ready_v13`:
  - АСД, Энтро, Хаусдорс, Олимп
- Один дубль Хаусдорс (`Под покраску / Под покраску`) удалён

## Финальные файлы (`output/final_hybrid_client_ready_v14/`)

- `doors_catalog.csv` — 3084 строки, 10 производителей
- `doors_catalog.xlsx` — XLSX-версия каталога
- `wordpress_import.csv` — 3084 строки (синхронизирован с каталогом)
- `qa_report.csv` — 177 линеек с разбивкой по строкам
- `run_report.json` — метрики сборки
- `not_found.csv` — 0 строк
- `unchanged_baseline_check.csv` — подтверждение, что 6 старых производителей не изменились
- `hybrid_merge_report.csv` — отчёт об источниках данных
- `check_by_producer/` — отдельные CSV/XLSX по каждому из 10 производителей

## Метрики финальной сборки (v14)

- total_rows: 3084
- producers_count: 10
- rows_with_image: 2997
- rows_without_image: 87
- needs_review: 362
- duplicates: 0
- old 6 producers exact match: true

## WordPress экспорт

`wordpress_import.csv` пересобирается строго из `doors_catalog.csv`.
Количество строк совпадает: 3084 = 3084.

## Запуск парсера

```powershell
python src/doors_parser.py --targets data/target_lines.csv --sites config/sites.yaml --out output/<name> --only "<producer>"
```
