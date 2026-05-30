# doors_parser_final

Проект собирает каталог дверей по целевым линейкам из `data/target_lines.csv`.

## Что входит в финальный пакет

- Парсер: `src/`
- Скрипты:
  - `scripts/merge_scambio_into_final.py` — объединение нового блока Скамбио с общим каталогом
  - `scripts/build_check_by_producer.py` — проверочные выгрузки по производителям
- Конфиг: `config/sites.yaml`
- Данные: `data/target_lines.csv`
- Финальный результат: `output/final_merged/`

## Ключевой сценарий сборки финала

1. Отдельно перепарсить Скамбио в `output/scambio_check/`.
2. Выполнить merge без перепарсинга остальных сайтов:

```powershell
python scripts/merge_scambio_into_final.py
```

Результат создаётся в `output/final_merged/` и не перезаписывает базовый `output/`.

## Финальные файлы (`output/final_merged/`)

- `doors_catalog.csv`
- `doors_catalog.xlsx`
- `wordpress_import.csv`
- `qa_report.csv`
- `not_found.csv`
- `image_audit.csv`
- `duplicates_report.csv`
- `run_report.json`
- `check_by_producer/` (отдельные выгрузки и аудит по 10 производителям)

## Актуальные метрики финальной сборки

По `output/final_merged/run_report.json`:

- `total_rows`: **3913**
- `producers_count`: **10**
- `target_lines_total`: **182**
- `target_lines_with_real_models`: **164**
- `target_lines_as_placeholder`: **18**
- `target_lines_not_found`: **0**
- `rows_with_image`: **3546**
- `rows_without_image`: **367**
- `material_filled_count`: **3507**
- `needs_review_count`: **1344**
- `duplicates_removed`: **0**
- `bad_images_removed`: **320**

## Скамбио (в финальном каталоге)

Источник замены: `output/scambio_check/doors_catalog.csv`.

В финальный каталог вставлен новый блок Скамбио:

- строк: **890**
- линеек: **21/21**
- присутствуют `ПСЛ`, `ПСЕ`, `Инвизибл`, все целевые `Смальта*`
- в `model_name` нет `Распродажа`, `размер`, `...`
- отсутствуют строки комплектующих

## WordPress экспорт

`output/final_merged/wordpress_import.csv` пересобирается строго из
`output/final_merged/doors_catalog.csv`.

Количество строк совпадает: **3913 = 3913**.

## Примечание по ручной проверке

Строки со статусом `needs_review` требуют визуальной проверки.
Для этого подготовлены:

- `output/final_merged/check_by_producer/*.xlsx`
- `output/final_merged/check_by_producer/audit_summary.csv`
- `output/final_merged/check_by_producer/samples_for_visual_check.xlsx`
