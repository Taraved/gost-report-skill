<div align="center">
  <h1>📑 GOST Report Skill</h1>
  <p><em>AI-agent skill for generating GOST 7.32-2001 compliant lab reports as .docx</em></p>
  <p>
    <img src="https://img.shields.io/badge/python-3.10%2B-blue" alt="Python 3.10+">
    <img src="https://img.shields.io/badge/license-MIT-green" alt="MIT License">
    <img src="https://img.shields.io/badge/status-beta-yellow" alt="Status: Beta">
  </p>
  <p>
    <a href="#-описание">Описание</a> •
    <a href="#-как-установить">Установка</a> •
    <a href="#-организация-файлов">Файлы лабы</a> •
    <a href="#-структура-проекта">Структура</a> •
    <a href="#-json-schema">JSON Schema</a>
  </p>
</div>

---

[🇷🇺 Русская версия](#русский) · [🇬🇧 English version](#english)

---

<a name="русский"></a>

# 🇷🇺 Русский

## 📋 Описание

**GOST Report Skill** — это инструкция (skill) для AI-агентов (Hermes Agent, Claude Code, Codex CLI и др.), которая позволяет автоматически оформлять отчёты по лабораторным работам в соответствии с **ГОСТ 7.32-2001**.

**Как это работает:**

1. Вы даёте агенту исходный код, скриншоты схем и, если есть, методичку.
2. Агент анализирует материалы (через Vision для схем, парсинг для кода).
3. Агент формирует структурированный `data.json` с описанием каждого задания.
4. Скрипт `generate_docx.py` превращает JSON в готовый `.docx`-файл с титульным листом, оглавлением, таблицами, листингами и нумерацией по ГОСТ.

Всё, что нужно от вас — дать агенту два файла и сказать: *«Сделай отчёт по лабораторной работе»*.

---

## 🤖 Как установить

### Hermes Agent

Положите `skill/SKILL.md` и `scripts/generate_docx.py` в директорию навыков:

```bash
cp skill/SKILL.md ~/.hermes/skills/gost-lab-docx/SKILL.md
cp scripts/generate_docx.py ~/.hermes/skills/gost-lab-docx/scripts/generate_docx.py
```

Затем скажите агенту: *«Используй навык gost-lab-docx для оформления отчёта»*.

Агент самостоятельно установит `python-docx`, проанализирует файлы, создаст `data.json` и запустит скрипт генерации.

### Claude Code

Скопируйте содержимое `skill/SKILL.md` в файл `CLAUDE.md` вашего проекта или укажите в системном промпте.

### Другие агенты

Любой агент с доступом к файловой системе и Python может:
1. Прочитать `skill/SKILL.md` как инструкцию.
2. Сформировать `data.json`.
3. Вызвать `python scripts/generate_docx.py`.

---

## 📂 Организация файлов для лучшего результата

Чтобы система не ошибалась, рекомендуется следующая структура в папке с вашей лабораторной работой:

### Одно задание

```text
Lab/
├── методичка.pdf   (опционально, для постановки задачи)
├── схема.jpeg      (скриншот схемы для анализа Vision)
├── main.cpp        (исходный код)
└── README.md       (ваши краткие заметки, если нужно)
```

### Несколько заданий

```text
Lab/
├── 1.jpeg          (скриншот схемы для Vision)
├── 1a.txt          (описание задания)
├── 1.txt           (исходный код)
├── 2.jpeg          (скриншот схемы для Vision)
├── 2a.txt          (описание задания)
├── 2.txt           (исходный код)
├── 3.jpeg          (скриншот схемы для Vision)
├── 3a.txt          (описание задания)
└── 3.txt           (исходный код)
```

---

## 📂 Структура проекта

```
gost-report-skill/
├── LICENSE                  # MIT лицензия
├── README.md                # Этот файл
├── pyproject.toml           # Python-пакет (PEP 621)
├── .gitignore
├── skill/
│   └── SKILL.md             # Инструкция для AI-агента
├── scripts/
│   ├── __init__.py
│   └── generate_docx.py     # Генератор .docx
├── examples/
│   ├── data.json            # Пример входных данных
│   ├── sample_report.docx   # Готовый отчёт-пример
│   └── sample_circuit.jpeg  # Пример схемы
└── docs/
    ├── screenshot_1.png
    ├── screenshot_2.png
    ├── screenshot_3.png
    ├── screenshot_4.png
    └── usage.md             # Справочник по ошибкам
```

---

## 📐 JSON Schema

Отчёт описывается JSON-файлом. Вот полная структура с пояснениями.

### Пример

```json
{
  "discipline": "Алгоритмы и техника программирования",
  "lab_number": 5,
  "student_name": "Иванов И.И.",
  "student_group": "РИ-231001",
  "teacher": "Петров П.П.",
  "city": "Екатеринбург",
  "year": "2026",
  "keywords": ["Qt", "C++", "GUI"],
  "introduction": "Цель работы — освоить базовые приёмы...",
  "conclusion": "Все задачи выполнены в полном объёме.",
  "tasks": [
    {
      "number": 1,
      "title": "Название задания",
      "problem_statement": "Текст постановки задачи",
      "test_data": {
        "headers": ["Ввод", "Вывод"],
        "rows": [["значение1", "результат1"]]
      },
      "algorithm": "1. Шаг первый...\n2. Шаг второй...",
      "code_files": [
        {
          "caption": "main.cpp",
          "content": "int main() { return 0; }"
        }
      ],
      "image_path": "docs/screenshot.png",
      "image_caption": "Схема подключения"
    }
  ],
  "bibliography": [
    "Автор. Название — М.: Изд-во, 2025. — 200 с."
  ],
  "appendix": [
    {
      "title": "Приложение А",
      "content": "Полный листинг..."
    }
  ]
}
```

### Обязательные поля

| Поле | Тип | Описание |
|------|-----|----------|
| `lab_number` | `int` | Номер лабораторной работы |
| `tasks` | `array` | Массив заданий (минимум 1) |

### Поля каждого задания (`tasks[]`)

| Поле | Тип | Обязательное | Описание |
|------|-----|:---:|----------|
| `number` | `int` | ✅ | Номер задания |
| `title` | `string` | ✅ | Краткое название |
| `problem_statement` | `string` | ✅ | Формальная постановка задачи |
| `test_data` | `object` | | Таблица с `headers` и `rows` |
| `algorithm` | `string` | | Пошаговое описание (через `\n`) |
| `code_files` | `array` | ✅ | Массив объектов `{caption, content}` |
| `image_path` | `string` | | Путь к скриншоту/схеме |
| `image_caption` | `string` | | Подпись к рисунку |

> **Примечание**: `test_data.headers` и `test_data.rows` не обязательны, но без них таблица в отчёте не появится.

---

## 📄 Лицензия

MIT — делайте что хотите, но упомяните автора.

---

## 🙌 Как помочь

- PR и Issues приветствуются.
- Нашли ошибку? [Создайте issue](https://github.com/Taraved/gost-report-skill/issues).
- Хотите добавить поддержку формул или сложных таблиц? Форк и вперёд.

---

<a name="english"></a>

# 🇬🇧 English

## 📋 Description

**GOST Report Skill** is an AI-agent skill (for Hermes Agent, Claude Code, Codex CLI, etc.) that automates lab report generation in **GOST 7.32-2001** format — the Russian state standard for academic documentation.

**How it works:**

1. You provide the agent with source code, screenshots, and (optionally) a lab manual.
2. The agent analyzes everything (Vision for diagrams, text parsing for code).
3. The agent produces a structured `data.json` describing each task.
4. The `generate_docx.py` script turns the JSON into a polished `.docx` report with title page, TOC, tables, listings, and GOST-compliant numbering.

All you need is to give the agent two files and say: *"Generate a lab report"*.

---

## 🤖 How to install

### Hermes Agent

Place `skill/SKILL.md` and `scripts/generate_docx.py` into your skills folder:

```bash
cp skill/SKILL.md ~/.hermes/skills/gost-lab-docx/SKILL.md
cp scripts/generate_docx.py ~/.hermes/skills/gost-lab-docx/scripts/generate_docx.py
```

Then tell your agent: *"Use the gost-lab-docx skill to format my lab report"*.

The agent will install `python-docx`, analyze your files, create `data.json`, and run the generation script — all automatically.

### Claude Code

Copy `skill/SKILL.md` contents into `CLAUDE.md` or your system prompt.

### Other agents

Any agent with filesystem and Python access can:
1. Read `skill/SKILL.md` as instructions.
2. Build `data.json`.
3. Run `python scripts/generate_docx.py`.

---

## 📂 File organization for best results

To help the agent understand your lab, structure your files like this:

### Single task

```text
Lab/
├── manual.pdf     (optional, for problem statement)
├── circuit.jpeg   (screenshot for Vision analysis)
├── main.cpp       (source code)
└── README.md      (your notes, if needed)
```

### Multiple tasks

```text
Lab/
├── 1.jpeg         (screenshot for Vision analysis)
├── 1a.txt         (task description)
├── 1.txt          (source code)
├── 2.jpeg         (screenshot for Vision analysis)
├── 2a.txt         (task description)
├── 2.txt          (source code)
├── 3.jpeg         (screenshot)
├── 3a.txt         (task description)
└── 3.txt          (source code)
```

---

## 📂 Project Structure

```
gost-report-skill/
├── LICENSE                  # MIT license
├── README.md                # This file
├── pyproject.toml           # Python packaging (PEP 621)
├── .gitignore
├── skill/
│   └── SKILL.md             # AI-agent instructions
├── scripts/
│   ├── __init__.py
│   └── generate_docx.py     # .docx generator
├── examples/
│   ├── data.json            # Sample input
│   ├── sample_report.docx   # Sample output
│   └── sample_circuit.jpeg  # Sample diagram
└── docs/
    ├── screenshot_1.png
    ├── screenshot_2.png
    ├── screenshot_3.png
    ├── screenshot_4.png
    └── usage.md             # Troubleshooting guide
```

---

## 📐 JSON Schema

The report is described by a JSON file. Here is the full structure:

### Example

```json
{
  "discipline": "Algorithms and Programming",
  "lab_number": 5,
  "student_name": "Ivanov I.I.",
  "student_group": "RI-231001",
  "teacher": "Petrov P.P.",
  "city": "Yekaterinburg",
  "year": "2026",
  "keywords": ["Qt", "C++", "GUI"],
  "introduction": "The goal of this work...",
  "conclusion": "All tasks completed.",
  "tasks": [
    {
      "number": 1,
      "title": "Task title",
      "problem_statement": "Formal problem description",
      "test_data": {
        "headers": ["Input", "Output"],
        "rows": [["value1", "result1"]]
      },
      "algorithm": "1. Step one...\n2. Step two...",
      "code_files": [
        {
          "caption": "main.cpp",
          "content": "int main() { return 0; }"
        }
      ],
      "image_path": "docs/screenshot.png",
      "image_caption": "Connection diagram"
    }
  ],
  "bibliography": ["Author. Title — Moscow: Publisher, 2025. — 200 p."],
  "appendix": [
    {
      "title": "Appendix A",
      "content": "Full listing..."
    }
  ]
}
```

### Required fields

| Field | Type | Description |
|-------|------|-------------|
| `lab_number` | `int` | Lab work number |
| `tasks` | `array` | Array of tasks (min 1) |

### Per-task fields (`tasks[]`)

| Field | Type | Required | Description |
|-------|------|:--------:|-------------|
| `number` | `int` | ✅ | Task number |
| `title` | `string` | ✅ | Short title |
| `problem_statement` | `string` | ✅ | Formal problem statement |
| `test_data` | `object` | | Table with `headers` and `rows` |
| `algorithm` | `string` | | Step-by-step description (use `\n`) |
| `code_files` | `array` | ✅ | Array of `{caption, content}` objects |
| `image_path` | `string` | | Path to screenshot/diagram |
| `image_caption` | `string` | | Image caption |

> **Note**: `test_data.headers` and `test_data.rows` are optional — without them no table is generated.

---

## 📄 License

MIT — do whatever you want, but give credit.

---

## 🙌 Contributing

PRs and issues are welcome. Found a bug? [Open an issue](https://github.com/Taraved/gost-report-skill/issues).
