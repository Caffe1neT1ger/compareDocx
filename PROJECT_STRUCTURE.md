# Структура проекта compareDocx

## Обзор

Проект организован следующим образом:

```
compareDocx/
├── .github/                    # GitHub Actions для CI/CD
│   └── workflows/
│       └── ci.yml              # Автоматическое тестирование
├── docs/                       # Документация проекта
│   ├── CONTENT_BASED_COMPARISON.md
│   ├── CUSTOM_STYLES_SUPPORT.md
│   ├── DOCUMENTATION.md        # Полная документация
│   ├── README.md               # Навигация по документации
│   └── TESTING_REPORT.md       # Отчет о тестировании
├── documents/                  # Тестовые документы
│   ├── test_document_1.docx
│   ├── test_document_2.docx
│   ├── extended_test_document_1.docx
│   ├── extended_test_document_2.docx
│   ├── additional_test_document_1.docx
│   ├── additional_test_document_2.docx
│   └── README.md
├── prompts/                    # Шаблоны промптов для LLM
│   ├── system_prompt.txt      # Системный промпт
│   ├── user_prompt_template.txt # Шаблон пользовательского промпта
│   └── README.md
├── tests/                      # Тесты проекта
│   ├── conftest.py            # Конфигурация pytest
│   ├── test_*.py              # Unit тесты
│   ├── comprehensive_test.py  # Комплексное тестирование
│   ├── edge_cases_test.py     # Тесты граничных случаев
│   ├── check_llm_results.py    # Скрипт проверки LLM результатов
│   ├── create_*_documents.py  # Скрипты генерации тестовых документов
│   ├── README.md
│   └── TEST_REPORT.md
├── *.py                        # Основной код проекта
├── .gitignore                  # Игнорируемые файлы
├── .gitattributes              # Настройки Git
├── README.md                   # Основная документация
├── requirements.txt            # Зависимости Python
└── env.example                 # Пример конфигурации .env
```

## Основные модули

### Парсинг документов
- **docx_file.py** - Парсинг DOCX документов, извлечение структуры, таблиц, изображений

### Сравнение
- **compare.py** - Логика сравнения документов, определение типов изменений

### Экспорт
- **excel_export.py** - Экспорт в Excel
- **json_export.py** - Экспорт в JSON
- **csv_export.py** - Экспорт в CSV
- **html_export.py** - Экспорт в HTML

### Интеграция LLM
- **llm_adapter.py** - Адаптер для работы с OpenAI LLM

### Утилиты
- **config.py** - Централизованная конфигурация
- **logger_config.py** - Настройка логирования
- **exceptions.py** - Кастомные исключения
- **validators.py** - Валидация входных данных

### Точки входа
- **cli.py** - Командная строка интерфейс
- **main.py** - Основной модуль (альтернативная точка входа)

## Игнорируемые файлы и папки

Следующие файлы и папки исключены из Git:

- `__pycache__/` - Кэш Python
- `*.pyc`, `*.pyo`, `*.pyd` - Скомпилированные файлы Python
- `.pytest_cache/` - Кэш pytest
- `.coverage`, `htmlcov/` - Отчеты о покрытии кода
- `.env` - Конфигурация с секретными данными
- `results/` - Результаты сравнения документов
- `*.xlsx`, `*.csv`, `*.json`, `*.html` - Выходные файлы (кроме тестов)
- `htmlcov/` - HTML отчеты о покрытии
- `docs/IMPROVEMENTS.md`, `docs/IMPROVEMENT_OPTIONS.md` - Временные файлы документации

## Тестирование

### Запуск тестов
```bash
# Все тесты
python -m pytest tests/ -v

# С покрытием кода
python -m pytest tests/ --cov=. --cov-report=html

# Конкретный тест
python -m pytest tests/test_compare.py -v
```

### Типы тестов
- **Unit тесты** - Тестирование отдельных функций и классов
- **Интеграционные тесты** - Тестирование полного цикла сравнения
- **Тесты граничных случаев** - Тестирование обработки ошибок

## Документация

- **README.md** - Основная документация и быстрый старт
- **docs/DOCUMENTATION.md** - Полная документация проекта
- **docs/CUSTOM_STYLES_SUPPORT.md** - Поддержка кастомных стилей
- **docs/CONTENT_BASED_COMPARISON.md** - Сравнение по содержимому
- **docs/TESTING_REPORT.md** - Отчет о тестировании

## Конфигурация

### Переменные окружения
Создайте файл `.env` на основе `env.example`:
```bash
OPENAI_API_KEY=your-api-key-here
OPENAI_API_URL=https://foundation-models.api.cloud.ru/v1
OPENAI_MODEL=openai/gpt-oss-120b
OPENAI_TEMPERATURE=0.3
OPENAI_MAX_TOKENS=200
```

