# Usage Guide — GOST Report Skill

## Быстрый запуск

```bash
# 1. Установить python-docx
pip install python-docx

# 2. Создать data.json (см. README или SKILL.md)
# 3. Запустить генерацию
python scripts/generate_docx.py -i data.json -o report.docx
```

## Проверка

```bash
# Проверить, что скрипт работает
python scripts/generate_docx.py -i examples/data.json -o /tmp/test_report.docx
ls -la /tmp/test_report.docx   # должен быть >0 байт
```

## Workflow для студента

1. Собрать все файлы лабы в одну папку: код, скриншоты, методичку.
2. Попросить AI-агента (с подключённым skill/SKILL.md) проанализировать папку.
3. Агент создаст `data.json` — проверить вручную.
4. Запустить `generate_docx.py`.
5. Открыть `.docx` в Word, нажать `Ctrl+A → F9` для обновления оглавления.
6. Сохранить как PDF, если требуется.

## Типичные ошибки

| Ошибка | Причина | Решение |
|--------|---------|---------|
| `ModuleNotFoundError: No module named 'docx'` | Не установлен python-docx | `pip install python-docx` |
| `FileNotFoundError` | Неверный путь в `--input` или `--output` | Проверить пути |
| `KeyError: 'tasks'` | В JSON нет поля `tasks` | Добавить `"tasks": [...]` |
| Пустое оглавление | Word не обновил поля | `Ctrl+A → F9` |
